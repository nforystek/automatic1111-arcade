@echo off
cd "C:\Development\Projects\Txt2ImgKiosk\webui"
set PYTHON=
set GIT=
set VENV_DIR=
Set SD_WEBUI_LOG_LEVEL = Info
set COMMANDLINE_ARGS= --xformers --no-prompt-history --theme dark --no-download-sd-model --do-not-download-clip--administrator --api --api-log --loglevel INFO --disable-tls-verify
call webui.bat
