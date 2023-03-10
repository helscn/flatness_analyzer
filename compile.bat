@echo off
SETLOCAL ENABLEEXTENSIONS

SET compile_name=Flatness
SET compile_path=%~dp0
SET env_path=%~dp0
SET compile_file=%compile_path%\%compile_name%.py


IF EXIST "%compile_path%\icon.ico" (
	SET package_icon=--icon "%compile_path%\icon.ico"
) ELSE (
	SET package_icon= 
)

REM SET compile_mode=--onefile
SET compile_mode=--onedir

REM SET show_console=--console
SET show_console=--noconsole

IF EXIST "%compile_path%\.env\Scripts\activate.bat" (
	ECHO 正在激活脚本运行的 Python 虚拟环境...
	SET env_path=%compile_path%\.env
	CALL "%compile_path%\.env\Scripts\activate.bat"
)
IF EXIST %compile_path%\.venv\Scripts\activate.bat (
	ECHO 正在激活脚本运行的 Python 虚拟环境...
	SET env_path=%compile_path%\.venv
	CALL "%compile_path%\.venv\Scripts\activate.bat"
)

ECHO.
ECHO PyInstaller %compile_mode% %show_console% %package_icon%  %version_file%  %compile_name% %compile_file%
pyinstaller %compile_mode% %show_console% %package_icon% --name %compile_name% "%compile_file%"
IF ERRORLEVEL 1 (
	ECHO.
	ECHO 打包 Python 脚本时出现错误! 
) ELSE (
	XCOPY %compile_path%\ui %compile_path%\dist\%compile_name%\ui /I /Q
	ECHO.
	ECHO  脚本 %compile_file% 打包完成。
)


IF EXIST "%compile_path%\file_version_info.txt" (
	IF EXIST "%env_path%\Lib\site-packages\PyInstaller\utils\cliutils\set_version.py" (
		ECHO 对打包程序写入文件版本信息...
		python "%env_path%\Lib\site-packages\PyInstaller\utils\cliutils\set_version.py" "%compile_path%\file_version_info.txt" "%compile_path%\dist\%compile_name%\%compile_name%.exe" 
	)
)

IF EXIST "%compile_path%\.env\Scripts\deactivate.bat" (
	ECHO.
	ECHO 正在退出脚本运行的 Python 虚拟环境...
	CALL "%compile_path%\.env\Scripts\deactivate.bat"
)
IF EXIST "%compile_path%\.venv\Scripts\deactivate.bat" (
	ECHO.
	ECHO 正在退出脚本运行的 Python 虚拟环境...
	CALL "%compile_path%\.venv\Scripts\deactivate.bat"
)
ECHO.
PAUSE
