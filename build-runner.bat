@echo off
setlocal

if "%~1"=="" (
  set CONFIG=Release
) else (
  set CONFIG=%~1
)

echo Building SharePoint2010Migration.Runner (%CONFIG%)...
msbuild Sharepoint2010.sln /t:Build /p:Configuration=%CONFIG% || goto :error

echo.
echo Executable generated at:
echo src\SharePoint2010Migration.Runner\bin\%CONFIG%\SharePoint2010Migration.Runner.exe
exit /b 0

:error
echo Build failed.
exit /b 1
