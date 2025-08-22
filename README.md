# AutoUp
浏览器自动上传文件

# nuitka 打包
```commandline
python -m nuitka --standalone --onefile --enable-plugin=tk-inter --windows-console-mode=disable --enable-plugin=playwright --playwright-include-browser=chromium-1169  .\main.py
```