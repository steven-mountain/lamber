@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
cd /d "d:\HermesJang\CMCC\tools\Benefitcostanalysis"
npx tauri dev
