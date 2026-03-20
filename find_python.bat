@echo off
REM Auto-find Python for clawTest

REM Check E:\Miniconda first
if exist "E:\Miniconda\python.exe" (
    echo E:\Miniconda\python.exe
    exit /b 0
)

REM Check D:\Miniconda
if exist "D:\Miniconda\python.exe" (
    echo D:\Miniconda\python.exe
    exit /b 0
)

REM Check D:\Miniconda3
if exist "D:\Miniconda3\python.exe" (
    echo D:\Miniconda3\python.exe
    exit /b 0
)

REM Check C:\Python
if exist "C:\Python\python.exe" (
    echo C:\Python\python.exe
    exit /b 0
)

REM Check user profile Python
if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python*\python.exe" (
    for /f "delims=" %%i in ('dir /b /ad "%USERPROFILE%\AppData\Local\Programs\Python\Python*" 2^>nul') do (
        if exist "%USERPROFILE%\AppData\Local\Programs\Python\%%i\python.exe" (
            echo %USERPROFILE%\AppData\Local\Programs\Python\%%i\python.exe
            exit /b 0
        )
    )
)

REM Check system PATH
where python >nul 2>&1
if %errorlevel% equ 0 (
    for /f "delims=" %%i in ('where python') do (
        echo %%i | findstr /i "WindowsApps" >nul
        if errorlevel 1 (
            echo %%i
            exit /b 0
        )
    )
)

echo ERROR
exit /b 1
