# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Environment

- This repository is a Windows-first local scanning toolkit, not a packaged application. The main workflows depend on bundled Windows executables (`ts.exe`, `spray.exe`), the vendored Python tool `tools/web_survivalscan/`, `.bat` wrappers, and Windows shell behavior used in `1.py` (`start cmd /c`, `ping -n`, `chcp 65001`).
- Run the Python entrypoints from the repository root. `2.py`, `2.txt`, and `ppp.py` use fixed filenames in the current working directory; `1.py` resolves paths relative to its own location.
- There is no configured build system, linter, or automated test suite in this repo.

## Common commands

### Install Python dependencies

```bash
python -m pip install pandas openpyxl psutil pywin32
python -m pip install -r tools/web_survivalscan/requirements.txt
```

- `pywin32` is only needed for the optional console-hiding path in `1.py`.
- Full workflows also require the repo-local binaries `ts.exe` and `spray.exe`.
- `tools/web_survivalscan/requirements.txt` adds the vendored Web-SurvivalScan dependencies (`termcolor`, `tqdm`, `requests`, `urllib3`, `aiofiles`, `httpx`, `beautifulsoup4`).

### Run the single-port ts URL flow

```bash
python 2.py
```

- Runs `ts -hf ip.txt -pa 3389 -np -m port,url`
- Parses `url.txt`
- Generates a styled `url_details_*.xlsx`
- Rewrites `url.txt` with cleaned URLs

### Run the ports-file ts URL flow

```bash
python 2.txt
```

- Despite the extension, `2.txt` is Python source code.
- Runs `ts -hf ip.txt -portf ports.txt -np -m port,url -t 100`
- Produces the same normalized `url.txt` / Excel outputs as `2.py`, but scans the port set from `ports.txt`

### Run the port/fingerprint report flow

```bash
python ppp.py
```

- Reads `port.txt`
- Normalizes mixed port-status, fingerprint, and URL records into one styled `port_scan_report_*.xlsx`

### Run the full spray → filter → Web-SurvivalScan pipeline

```bash
python 1.py
```

- Runs `spray.exe -l url.txt -d dirv2.txt -f res.json`
- Post-processes spray output via `process_data.py`
- Filters status-200 URLs into a dated text file
- Runs the vendored `tools/web_survivalscan/Web-SurvivalScan.py` non-interactively against that filtered URL list
- Archives spray outputs plus Web-SurvivalScan artifacts (`output.txt`, `outerror.txt`, `.data/report.json`, `report.html`) into a `MMDD/` folder

### Run focused post-processing on one artifact

```bash
python process_data.py <input.json|input.xlsx> <output.xlsx>
```

- `.json` input: treat as spray JSONL output and create a filtered Excel report plus a sibling `.txt` URL list
- `.xlsx`/`.xls` input: legacy ehole workbook beautification path; current primary workflow no longer depends on ehole

### Batch wrappers

- `端口处理.bat`: runs `python ppp.py`
- `top100仅端口.bat`: runs `python 2.py`, then `python ppp.py`
- `top1000仅端口.bat`: runs `python 2.txt`, then `python ppp.py`
- `轮子top100.bat` / `轮子top1000.bat`: run the ts preprocessing step (`2.py` or `2.txt`), run `ppp.py`, clean `res.json` and `res_processed.*`, then start `python 1.py`
- `小字典.bat`: runs `python 1.py`

## Manual verification

- There is no unit-test or single-test runner in this repo.
- Use the smallest relevant script as a manual verification step:
  - `python process_data.py <artifact> <output.xlsx>` for post-processing logic only
  - `python ppp.py` for `port.txt` parsing and report generation
  - `python 2.py` or `python 2.txt` for ts URL parsing and Excel generation
  - `python 1.py` for the full spray → filter → Web-SurvivalScan pipeline
  - `python tools/web_survivalscan/Web-SurvivalScan.py` for the vendored survival checker by itself
- The Python scripts are written as top-level entrypoints rather than importable modules, so “single test” in this repo means running the narrowest script that covers the code path you changed.

## High-level architecture

- The repository is organized as a pipeline around external scanners, with Python handling normalization, orchestration, and Excel reporting rather than doing the scanning itself.
- `2.py` and `2.txt` are the first-stage ts preprocessing entrypoints. Both run `ts`, parse `url.txt`, generate a styled `url_details_*.xlsx`, and then rewrite `url.txt` with de-duplicated cleaned URLs for downstream consumers. Their only material difference is the ts invocation (`-pa 3389` in `2.py` vs `-portf ports.txt -t 100` in `2.txt`).
- `ppp.py` is a parallel reporting path for `port.txt`. It parses mixed line formats from ts output—plain port-state rows, fingerprint rows, and URL rows—and merges them into one `port_scan_report_*.xlsx` workbook with conditional formatting.
- `1.py` is the top-level Windows orchestration layer for the spray → process_data → Web-SurvivalScan flow. It runs external commands, streams their output to timestamped log files, monitors process completion and progress files, creates a date folder, filters 200-status URLs, invokes the vendored `tools/web_survivalscan/Web-SurvivalScan.py` non-interactively, and moves finished artifacts into `MMDD/` output directories.
- `process_data.py` is the shared normalization/reporting layer used by `1.py`. For spray JSONL it preserves the original rows in Excel, strips `redirect_url`, flattens selected nested fingerprint fields, emits a sibling URL list, and leaves the 200-only filtering decision to `1.py`. Its Excel-input branch is now a legacy ehole-beautification path rather than part of the primary workflow.
- Column matching between `1.py` and `process_data.py` is intentionally loose: both use candidate header-name lists plus column-index fallbacks so downstream filtering can survive schema variations in external tool output.
- The `.bat` files are operator wrappers around the Python entrypoints. The `轮子top100*.bat` wrappers are destructive to intermediate state: they delete local `port.txt`, `url.txt`, `res.json`, and `res_processed.*` before or between stages, then launch the full pipeline.
- Files such as `ip.txt`, `ports.txt`, `dirv2.txt`, `bak.txt`, `bak2.txt`, `finger.json`, `config.yaml`, and `TscanClient.txt` are runtime inputs, dictionaries, or tool/config artifacts rather than importable Python modules.
- Root-level intermediate outputs include `port.txt`, `url.txt`, `res.json`, `res_processed.xlsx`, and `res_processed.txt`. Finalized spray artifacts and Web-SurvivalScan outputs are archived into date folders named `MMDD/`.

## Additional documentation

- `docs/superpowers/specs/2026-04-19-ppp-domain-support-design.md` records the intended compatibility boundary for `ppp.py` host parsing: domain names are supported in `port.txt`, IPv6 is still out of scope, and the Excel schema stays unchanged.
- `docs/superpowers/plans/2026-04-19-ppp-domain-support.md` is the implementation plan for that parser change and is useful historical context when adjusting `ppp.py`.

## Repo-specific notes

- `1.py` is explicitly written for Windows shell semantics (`start cmd /c`, `ping -n`, `chcp 65001`). Do not replace those with POSIX equivalents unless the task is to port the workflow.
- Run workflow scripts from the repo root, because several entrypoints use fixed filenames in the current working directory instead of resolving all paths relative to imports or arguments.
- The scripts assume fixed filenames and in-place cleanup.
- `process_data.py` is the canonical place for spray output shaping and workbook beautification. If a change affects spray Excel structure or URL extraction, update shared logic there rather than forking formatting logic in `1.py`.
- `ppp.py` accepts either IPv4 addresses or domain-style hosts in `port.txt` records, but still writes the target value into the `IP地址` column. Preserve that schema unless the task explicitly requires a report format change.
- `ppp.py` does not support IPv6 bare hosts in `port.txt`; keep that boundary in mind when changing its regex-based parser.
- Do not casually rename workbook headers or output columns. `1.py`, `process_data.py`, and downstream manual workflows rely on stable header names with loose matching rather than a formal schema layer.
- The `.bat` wrappers, `2.py` / `2.txt`, and `1.py` all overwrite or clean up root-level working files such as `port.txt`, `url.txt`, `res.json`, `res_processed.xlsx`, and `res_processed.txt`. Treat those scripts as stateful operators, not side-effect-free test commands.
- `config.yaml` appears to be external tool configuration/state rather than configuration loaded by the Python scripts directly.
