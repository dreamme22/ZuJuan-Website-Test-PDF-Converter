@echo off
rem ============================================
rem  PDF 一键转换：拖入 PDF 到此 bat 上即可
rem  或双击后按提示输入 PDF 路径
rem ============================================

rem 优先使用已装好依赖的 Python；不存在则退回 PATH 里的 python
set "SDKPY=C:\Users\19618\python-sdk\python3.13.2\python.exe"
set "PY=%SDKPY%"
if not exist "%SDKPY%" set "PY=python"

set "PDFFILE=%~1"
if "%PDFFILE%"=="" set /p PDFFILE=请输入 PDF 文件路径，可直接拖入文件：

"%PY%" "%~dp0cnvpdf.py" "%PDFFILE%"

echo.
pause
