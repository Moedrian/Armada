@echo off

call .\build_ArmadaPack.bat

set oDir=.\artifacts\Armada
set pyDir=py_scripts
set pyFile=%pyDir%\pythonProject\main.py
set oDirPy=%oDir%\%pyDir%

if exist %oDirPy% (
    rmdir /s /q %oDirPy%
)

mkdir %oDirPy%
echo copying python file...
copy .\%pyFile% %oDirPy%
echo =======================================

echo Building New Project...
dotnet build .\src\NewProject -o %oDir%
echo =======================================