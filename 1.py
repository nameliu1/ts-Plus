import os
import time
import datetime
import sys
import json
import psutil
import subprocess
import shutil
import threading
import pandas as pd

# 配置信息
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(BASE_DIR, "url.txt")
DIR_FILE = os.path.join(BASE_DIR, "dirv2.txt")
JSON_FILE = os.path.join(BASE_DIR, "res.json")        # spray原始输出
STAT_FILE = os.path.join(BASE_DIR, "url.txt.stat")
HIDE_PYTHON_CONSOLE = False
MONITOR_INTERVAL = 5  # 进程监控间隔（秒）
STATUS_CODE_COL_INDEX = 9  # 兼容旧格式的J列索引
URL_COL_INDEX = 4  # 兼容旧格式的E列索引
STATUS_CODE_CANDIDATES = ["status", "状态码", "status_code", "code", "Status", "STATUS", "J"]
URL_CANDIDATES = ["url", "URL", "网址", "链接", "directurl", "direct_url", "Direct URL", "E"]
EHOLE_QUICK_TIMEOUT = 3  # ehole快速完成的超时时间（秒）

# 需要删除的过程文件列表
TO_DELETE_FILES = [
    os.path.join(BASE_DIR, "res_processed.txt")
]

def log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def hide_python_console():
    if HIDE_PYTHON_CONSOLE:
        try:
            import win32gui, win32con
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        except:
            log("警告: 无法隐藏 Python 控制台窗口")

def _stream_process_output(process, log_handle):
    try:
        for line in iter(process.stdout.readline, ''):
            if not line:
                break
            print(line, end='')
            log_handle.write(line)
            log_handle.flush()
    finally:
        if process.stdout:
            process.stdout.close()


def run_native_command(command, process_name):
    command_text = subprocess.list2cmdline(command) if isinstance(command, list) else command
    log(f"执行命令: {command_text}")

    # 生成日志文件路径
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(BASE_DIR, f"{process_name}_{timestamp}.log")

    creationflags = 0
    if os.name == 'nt' and HIDE_PYTHON_CONSOLE:
        creationflags = subprocess.CREATE_NO_WINDOW

    log_handle = open(log_file, 'w', encoding='utf-8', errors='ignore')
    process = subprocess.Popen(
        command,
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        shell=isinstance(command, str),
        creationflags=creationflags,
        text=True,
        encoding='utf-8',
        errors='ignore',
        bufsize=1
    )
    output_thread = threading.Thread(
        target=_stream_process_output,
        args=(process, log_handle),
        daemon=True
    )
    output_thread.start()

    process.log_file = log_file
    process.log_handle = log_handle
    process.output_thread = output_thread
    log(f"命令输出将保存到: {log_file}")
    log(f"已启动进程: {process_name} (PID: {process.pid})")
    return process

def monitor_process(process_name, process=None, timeout=3600, progress_file=None, stat_file=None):
    log(f"监控进程: {process_name}")
    start_time = time.time()

    if process is not None:
        last_report_time = 0
        last_progress_size = -1
        stagnant = False

        while time.time() - start_time < timeout:
            return_code = process.poll()
            if return_code is not None:
                print()  # 进度条换行
                if hasattr(process, 'output_thread') and process.output_thread:
                    process.output_thread.join(timeout=2)
                if hasattr(process, 'log_handle') and process.log_handle and not process.log_handle.closed:
                    process.log_handle.close()
                log(f"进程已结束: {process_name} (退出码: {return_code})")
                if return_code != 0 and hasattr(process, 'log_file'):
                    log(f"警告: {process_name} 执行失败，请查看日志: {process.log_file}")
                return return_code == 0

            now = time.time()
            current_size = None
            if progress_file and os.path.exists(progress_file):
                current_size = os.path.getsize(progress_file)
                stagnant = current_size == last_progress_size
                last_progress_size = current_size
            else:
                stagnant = False

            stat_summary = ""
            if stat_file and os.path.exists(stat_file):
                try:
                    with open(stat_file, 'r', encoding='utf-8', errors='ignore') as f:
                        stat_content = f.read().strip()
                    if stat_content:
                        stat_data = json.loads(stat_content)
                        current_url = stat_data.get('url')
                        end_count = stat_data.get('end')
                        total_count = stat_data.get('total')
                        req_total = stat_data.get('req_total')
                        found = stat_data.get('found')
                        check = stat_data.get('check')
                        if current_url:
                            stat_summary += f" | 当前: {current_url}"
                        if end_count is not None and total_count is not None and total_count > 0:
                            pct = int(end_count * 100 / total_count)
                            filled = int(pct / 5)
                            bar = '█' * filled + '░' * (20 - filled)
                            stat_summary += f" | 字典: [{bar}] {pct}% ({end_count}/{total_count})"
                        if req_total is not None:
                            stat_summary += f" | 已发请求: {req_total}"
                        if found is not None:
                            stat_summary += f" | 发现: {found}"
                        if check is not None:
                            stat_summary += f" | 校验: {check}"
                except Exception:
                    pass

            if now - last_report_time >= MONITOR_INTERVAL:
                elapsed = int(now - start_time)
                if stat_summary:
                    print(f"\r{process_name} 运行中 {elapsed}s{stat_summary}", end='', flush=True)
                elif progress_file and os.path.exists(progress_file):
                    size_mb = current_size / 1048576 if current_size else 0
                    print(f"\r{process_name} 运行中 {elapsed}s | 输出: {size_mb:.1f}MB", end='', flush=True)
                else:
                    print(f"\r{process_name} 运行中 {elapsed}s", end='', flush=True)
                last_report_time = now

            time.sleep(1)

        process.kill()
        if hasattr(process, 'output_thread') and process.output_thread:
            process.output_thread.join(timeout=2)
        if hasattr(process, 'log_handle') and process.log_handle and not process.log_handle.closed:
            process.log_handle.close()
        log(f"错误: {process_name} 运行超时")
        if hasattr(process, 'log_file'):
            log(f"请查看日志文件: {process.log_file}")
        return False

    # 兼容旧逻辑：按进程名监控
    # 特殊处理ehole进程的快速完成情况
    is_ehole = process_name.lower() == "ehole.exe"
    quick_timeout = EHOLE_QUICK_TIMEOUT

    # 等待进程启动
    while time.time() - start_time < timeout:
        if any(proc.name().lower() == process_name.lower() for proc in psutil.process_iter()):
            log(f"进程已启动: {process_name}")
            break

        # 检查是否是ehole并且已经超过快速超时时间
        if is_ehole and (time.time() - start_time > quick_timeout):
            log(f"警告: ehole在{quick_timeout}秒内未启动，可能已快速完成")
            return True

        time.sleep(1)
    else:
        log(f"错误: 等待 {process_name} 启动超时")
        return False

    start_time = time.time()
    # 等待进程结束
    while time.time() - start_time < timeout:
        if not any(proc.name().lower() == process_name.lower() for proc in psutil.process_iter()):
            log(f"进程已结束: {process_name}")
            return True
        time.sleep(1)
    log(f"错误: {process_name} 运行超时")
    return False

def wait_for_file(file_path, timeout=300, require_non_empty=False):
    log(f"等待文件生成: {file_path}")
    start_time = time.time()
    while time.time() - start_time < timeout:
        if os.path.exists(file_path):
            if require_non_empty and os.path.getsize(file_path) == 0:
                time.sleep(1)
                continue
            log(f"文件已生成: {file_path}")
            return True
        time.sleep(1)
    log(f"错误: 文件未生成: {file_path}")
    return False

# 生成不冲突的文件名
def generate_unique_filename(base_dir, base_name, ext):
    counter = 1
    original_name = f"{base_name}{ext}"
    full_path = os.path.join(base_dir, original_name)
    
    # 如果文件已存在，则添加序号后缀
    while os.path.exists(full_path):
        new_name = f"{base_name}_{counter}{ext}"
        full_path = os.path.join(base_dir, new_name)
        counter += 1
    
    return full_path

# 删除指定的过程文件
def clean_process_files():
    log("开始清理上次运行的过程文件...")
    for file_path in TO_DELETE_FILES:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                log(f"已删除: {file_path}")
            except Exception as e:
                log(f"删除文件 {file_path} 时出错: {e}")
        else:
            log(f"文件不存在，跳过删除: {file_path}")
    log("过程文件清理完成")

def process_spray_output(json_file, excel_file):
    log(f"开始处理spray结果: {json_file}")
    process_data_script = os.path.join(BASE_DIR, "process_data.py")
    result = subprocess.run(
        ["python", process_data_script, json_file, excel_file],
        capture_output=True,
        text=True,
        cwd=BASE_DIR
    )
    if result.stdout.strip():
        log(f"process_data.py 输出:\n{result.stdout.strip()}")
    if result.stderr.strip():
        log(f"process_data.py 错误输出:\n{result.stderr.strip()}")
    if result.returncode != 0:
        log(f"错误: 数据处理失败，返回码: {result.returncode}")
        return None

    txt_file = os.path.splitext(excel_file)[0] + ".txt"

    if not os.path.exists(excel_file) or os.path.getsize(excel_file) == 0:
        log(f"错误: 处理后的Excel文件未生成或为空: {excel_file}")
        return None
    if not os.path.exists(txt_file):
        log(f"错误: 未找到URL列表文件: {txt_file}")
        return None

    with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
        url_count = len([line for line in f.readlines() if line.strip()])
    log(f"成功提取 {url_count} 个URL")
    return {"excel_file": excel_file, "txt_file": txt_file, "url_count": url_count}

def _normalize_column_name(value):
    return str(value).strip().lower()


def _find_column(df, candidates, fallback_index=None):
    normalized_map = {}
    for column in df.columns:
        normalized_map.setdefault(_normalize_column_name(column), column)

    for candidate in candidates:
        matched = normalized_map.get(_normalize_column_name(candidate))
        if matched is not None:
            return matched

    if fallback_index is not None and len(df.columns) > fallback_index:
        fallback_column = df.columns[fallback_index]
        log(f"警告: 未命中候选列名，回退使用第 {fallback_index + 1} 列: {fallback_column}")
        return fallback_column
    return None


def filter_status_200(excel_file, output_dir, count):
    try:
        log(f"开始从 {excel_file} 中筛选状态码为200的URL...")
        if not os.path.exists(excel_file):
            log(f"错误: Excel文件不存在: {excel_file}")
            return {"success": False, "reason": "excel_missing"}

        df = pd.read_excel(excel_file)
        if df.empty:
            log("错误: Excel文件为空")
            return {"success": False, "reason": "excel_empty"}

        status_code_col = _find_column(df, STATUS_CODE_CANDIDATES, STATUS_CODE_COL_INDEX)
        url_col = _find_column(df, URL_CANDIDATES, URL_COL_INDEX)

        if status_code_col is None or url_col is None:
            log("错误: 未找到状态码列或URL列")
            log(f"Excel实际列数: {len(df.columns)}，列名: {list(df.columns)}")
            return {"success": False, "reason": "missing_columns"}

        log(f"使用列 '{url_col}' 作为URL列，列 '{status_code_col}' 作为状态码列")

        df[status_code_col] = pd.to_numeric(df[status_code_col], errors='coerce')
        df_200 = df[(df[status_code_col] == 200) & (df[url_col].notna())].copy()
        total_rows = len(df)
        filtered_rows = len(df_200)
        log(f"Excel总行数: {total_rows}，状态码为200的行数: {filtered_rows}")

        if filtered_rows == 0:
            log("警告: 未找到状态码为200的URL，本次将跳过ehole阶段")
            return {"success": True, "has_results": False, "output_file": None, "count": 0}

        urls_200 = df_200[url_col].astype(str).str.strip()
        urls_200 = urls_200[urls_200 != ""].drop_duplicates().tolist()
        log(f"提取并去重后得到 {len(urls_200)} 个状态码为200的URL")

        if not urls_200:
            log("警告: 200状态码记录存在，但URL列为空，本次将跳过ehole阶段")
            return {"success": True, "has_results": False, "output_file": None, "count": 0}

        date_str = datetime.datetime.now().strftime("%Y%m%d")
        base_filename = f"{date_str}_status200_urls_{count}"
        output_file = generate_unique_filename(output_dir, base_filename, ".txt")

        log(f"将状态码为200的URL写入文件: {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(urls_200))

        with open(output_file, 'r', encoding='utf-8') as f:
            written_urls = f.read().splitlines()

        if len(written_urls) != len(urls_200):
            log(f"警告: 写入的URL数量({len(written_urls)})与筛选的URL数量({len(urls_200)})不一致")

        log(f"状态码为200的URL已保存至: {output_file}")
        return {"success": True, "has_results": True, "output_file": output_file, "count": len(urls_200)}
    except Exception as e:
        log(f"筛选错误: {e}")
        return {"success": False, "reason": "exception"}

def main():
    try:
        hide_python_console()
        log(f"开始自动化漏洞扫描和指纹识别流程")
        log(f"基础目录: {BASE_DIR}")
        
        # 创建日期文件夹
        date_folder = datetime.datetime.now().strftime("%m%d")
        full_date_dir = os.path.join(BASE_DIR, date_folder)
        os.makedirs(full_date_dir, exist_ok=True)
        log(f"创建日期文件夹: {full_date_dir}")
        
        # 清理指定的过程文件
        clean_process_files()
        
        # 步骤1: 执行spray扫描
        log("步骤1: 执行spray扫描...")
        spray_cmd = f'spray.exe -l "{URL_FILE}" -d "{DIR_FILE}" -f "{JSON_FILE}"'
        spray_process = run_native_command(spray_cmd, "spray.exe")
        if not monitor_process("spray.exe", process=spray_process, timeout=1800, progress_file=JSON_FILE, stat_file=STAT_FILE):
            log("错误: spray执行失败或超时")
            sys.exit(1)
        if not wait_for_file(JSON_FILE):
            log("错误: spray未生成结果文件")
            sys.exit(1)
        
        # 步骤2: 处理spray结果，提取有效URL
        log("步骤2: 处理spray结果，提取有效URL...")

        unique_excel_file = generate_unique_filename(BASE_DIR, "res_processed", ".xlsx")
        spray_output = process_spray_output(JSON_FILE, unique_excel_file)
        if not spray_output:
            log("错误: 处理spray输出失败")
            sys.exit(1)

        # 步骤3: 筛选状态码200的URL
        log("步骤3: 筛选状态码200的URL...")
        filter_result = filter_status_200(spray_output["excel_file"], full_date_dir, 1)
        if not filter_result.get("success"):
            log("错误: 状态码200筛选失败")
            sys.exit(1)

        filtered_txt_path = filter_result.get("output_file")
        
        # 步骤3.5: 移动Spray结果文件到日期文件夹
        log("步骤3.5: 移动Spray结果文件到日期文件夹...")
        
        # 为移动的文件生成唯一文件名
        spray_json_base = f"spray_original_{datetime.datetime.now().strftime('%Y%m%d')}"
        spray_json_dest = generate_unique_filename(full_date_dir, spray_json_base, ".json")
        
        spray_excel_base = f"spray_processed_{datetime.datetime.now().strftime('%Y%m%d')}"
        spray_excel_dest = generate_unique_filename(full_date_dir, spray_excel_base, ".xlsx")
        
        shutil.move(JSON_FILE, spray_json_dest)
        log(f"已移动Spray原始结果: {spray_json_dest}")
        
        shutil.move(spray_output["excel_file"], spray_excel_dest)
        log(f"已移动Spray处理后Excel: {spray_excel_dest}")

        spray_txt_source = spray_output["txt_file"]
        if os.path.exists(spray_txt_source):
            spray_txt_base = f"spray_urls_{datetime.datetime.now().strftime('%Y%m%d')}"
            spray_txt_dest = generate_unique_filename(full_date_dir, spray_txt_base, ".txt")
            shutil.move(spray_txt_source, spray_txt_dest)
            log(f"已移动Spray提取URL列表: {spray_txt_dest}")

        if not filter_result.get("has_results"):
            log(f"自动化流程完成：Spray阶段已完成，但未发现状态码200的URL，已跳过ehole。结果保存在: {full_date_dir}")
            return

        if not filtered_txt_path or not os.path.exists(filtered_txt_path) or os.path.getsize(filtered_txt_path) == 0:
            log("错误: ehole输入文件不存在或为空")
            sys.exit(1)

        # 步骤4: 执行ehole指纹识别
        log("步骤4: 执行ehole指纹识别...")

        ehole_base = f"ehole_result_{datetime.datetime.now().strftime('%Y%m%d')}"
        ehole_output = generate_unique_filename(full_date_dir, ehole_base, ".xlsx")

        ehole_cmd = f'ehole finger -l "{filtered_txt_path}" -o "{ehole_output}" -t 10'
        ehole_process = run_native_command(ehole_cmd, "ehole.exe")

        if not monitor_process("ehole.exe", process=ehole_process, timeout=1800):
            log("错误: ehole执行失败或超时")

        if not wait_for_file(ehole_output, require_non_empty=True):
            log("错误: ehole未生成结果文件")
            sys.exit(1)

        log("美化ehole结果表格...")
        process_data_script = os.path.join(BASE_DIR, "process_data.py")
        beautify_result = subprocess.run(
            ["python", process_data_script, ehole_output, ehole_output],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
        if beautify_result.stdout.strip():
            log(f"ehole美化输出:\n{beautify_result.stdout.strip()}")
        if beautify_result.stderr.strip():
            log(f"ehole美化错误输出:\n{beautify_result.stderr.strip()}")
        if beautify_result.returncode != 0:
            log(f"警告: ehole结果表格美化失败，返回码: {beautify_result.returncode}")
        else:
            log("ehole结果表格美化完成")

        log(f"自动化流程全部完成！所有结果保存在: {full_date_dir}")

    except Exception as e:
        log(f"程序异常: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    os.system("chcp 65001 >nul 2>&1")  # 确保中文显示正常
    
    # 检查依赖
    try:
        import psutil
        import pandas as pd
    except ImportError:
        log("错误: 缺少psutil或pandas库，请执行 'pip install psutil pandas'")
        sys.exit(1)
    
    main()