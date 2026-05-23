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
