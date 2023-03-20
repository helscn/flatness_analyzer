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
	ECHO [32mÕýÔÚ¼¤»î½Å±¾ÔËÐÐµÄ Python ÐéÄâ»·¾³...[0m
	SET env_path=%compile_path%\.env
	CALL "%compile_path%\.env\Scripts\activate.bat"
)
IF EXIST %compile_path%\.venv\Scripts\activate.bat (
	ECHO [32mÕýÔÚ¼¤»î½Å±¾ÔËÐÐµÄ Python ÐéÄâ»·¾³...[0m
	SET env_path=%compile_path%\.venv
	CALL "%compile_path%\.venv\Scripts\activate.bat"
)

ECHO.
ECHO [31mPyInstaller %compile_mode% %show_console% %version_file% %package_icon% %add_data% --name %compile_name% "%compile_file%"
ECHO [90m
pyinstaller %compile_mode% %show_console% %version_file% %package_icon% %add_data% --name %compile_name% "%compile_file%"
IF ERRORLEVEL 1 (
	ECHO.
	ECHO [31m´ò°ü Python ½Å±¾Ê±³öÏÖ´íÎó! [0m
) ELSE (
	ECHO.
	ECHO [32m½Å±¾ %compile_file% ´ò°üÍê³É¡£[0m
)

IF EXIST "%compile_path%\.env\Scripts\deactivate.bat" (
	ECHO.
	ECHO [32mÕýÔÚÍË³ö½Å±¾ÔËÐÐµÄ Python ÐéÄâ»·¾³...[0m
	CALL "%compile_path%\.env\Scripts\deactivate.bat"
)
IF EXIST "%compile_path%\.venv\Scripts\deactivate.bat" (
	ECHO.
	ECHO [32mÕýÔÚÍË³ö½Å±¾ÔËÐÐµÄ Python ÐéÄâ»·¾³...[0m
	CALL "%compile_path%\.venv\Scripts\deactivate.bat"
)
ECHO.
PAUSE
