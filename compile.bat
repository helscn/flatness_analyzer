@echo off
SETLOCAL ENABLEEXTENSIONS

SET compile_name=Flatness

SET add_data=--add-data "ui\*.*;.\ui" --add-data ".\changelog.md;."

REM SET compile_mode=--onefile
SET compile_mode=--onedir

REM SET show_console=--console
SET show_console=--noconsole


SET compile_path=%~dp0
SET env_path=%~dp0
SET compile_file=%compile_path%%compile_name%.py

IF EXIST "%compile_path%icon.ico" (
	SET package_icon=--icon "%compile_path%icon.ico"
) ELSE (
	SET package_icon= 
)

IF EXIST "%compile_path%\file_version_info.txt" (
	SET version_file=--version-file "%compile_path\%file_version_info.txt"
) ELSE (
	SET version_file= 
)

ECHO.
IF EXIST "%compile_path%\.env\Scripts\activate.bat" (
	ECHO [32m���ڼ���ű����е� Python ���⻷��...[0m
	SET env_path=%compile_path%\.env
	CALL "%compile_path%\.env\Scripts\activate.bat"
)
IF EXIST %compile_path%\.venv\Scripts\activate.bat (
	ECHO [32m���ڼ���ű����е� Python ���⻷��...[0m
	SET env_path=%compile_path%\.venv
	CALL "%compile_path%\.venv\Scripts\activate.bat"
)

ECHO.
ECHO [31mPyInstaller %compile_mode% %show_console% %version_file% %package_icon% %add_data% --name %compile_name% "%compile_file%"
ECHO [90m
pyinstaller %compile_mode% %show_console% %version_file% %package_icon% %add_data% --name %compile_name% "%compile_file%"
IF ERRORLEVEL 1 (
	ECHO.
	ECHO [31m��� Python �ű�ʱ���ִ���! [0m
) ELSE (
	ECHO.
	ECHO [32m�ű� %compile_file% �����ɡ�[0m
)

IF EXIST "%compile_path%\.env\Scripts\deactivate.bat" (
	ECHO.
	ECHO [32m�����˳��ű����е� Python ���⻷��...[0m
	CALL "%compile_path%\.env\Scripts\deactivate.bat"
)
IF EXIST "%compile_path%\.venv\Scripts\deactivate.bat" (
	ECHO.
	ECHO [32m�����˳��ű����е� Python ���⻷��...[0m
	CALL "%compile_path%\.venv\Scripts\deactivate.bat"
)
ECHO.
PAUSE
