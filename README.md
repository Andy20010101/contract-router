# Contract Router

Windows 10 打包建议使用 Python 3.12。

## 打包 exe

在 Windows 10 或更新系统中安装 Python 3.12，然后在本项目目录双击或运行：

```bat
build_windows.bat
```

成功后可执行文件在：

```text
dist\contract-router.exe
```

## 运行源码

```bat
py -3.12 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python file_classifier.py
```

程序首次启动会在程序同目录生成 `app_settings.json`，并使用 `data` 目录保存默认台账路径。

## Shell E2E test-mode

源码或 Windows exe 都可以从命令行运行同一套无 GUI 点击的 E2E：

```bat
python file_classifier.py e2e --input tmp\e2e\source\input --output tmp\e2e\source\output --report tmp\e2e\source\report.json --log tmp\e2e\source\logs\e2e.log
dist\contract-router.exe e2e --input tmp\e2e\exe\input --output tmp\e2e\exe\output --report tmp\e2e\exe\report.json --log tmp\e2e\exe\logs\e2e.log
```

`report.json` 使用 `contract-router-e2e/v1` schema，失败时返回非零 exit code。内置 fixture 覆盖：

- 乱命名文件仍按内容字段和父目录发票号归入同一套退税材料。
- `LY25117260` / `25117260` 标准化为同一票。
- 报关单号、发票号、金额、日期、目的国形成闭环证据。
- 金额冲突、重复发票材料、无关文件分别在 report 中标记 `conflict`、`duplicate`、`unmatched`。

## Windows SSH E2E

项目根目录提供 `.win-e2e.yaml`，供 `win-ssh-e2e-runner` 从 Mac/Linux 通过 SSH 在 Windows 原生 PowerShell 中执行：

```bash
python /path/to/win-ssh-e2e-runner/runner/win_e2e.py doctor --host win-native
python /path/to/win-ssh-e2e-runner/runner/win_e2e.py run --config .win-e2e.yaml
```

当前配置使用 `sync.mode: manual`，适合在未提交的本地改动需要先打包同步到 Windows 时使用。流程会在 Windows 上安装依赖、运行 pytest、运行源码 test-mode、构建 `dist\contract-router.exe`，再用 exe 运行同一套 E2E，并拉回 `tmp\e2e`、日志和 exe 产物。
