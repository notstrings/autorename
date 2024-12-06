@echo off
pushd %~dp0
powershell -ExecutionPolicy Bypass -File ".\autorename.ps1" %*
popd
