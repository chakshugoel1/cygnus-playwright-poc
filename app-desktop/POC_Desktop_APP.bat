@echo off
setlocal
set "APP_ROOT=%~dp0"
pushd "%APP_ROOT%"
"%APP_ROOT%node_modules\electron\dist\electron.exe" .
popd
