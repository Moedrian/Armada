@echo off

call .\build_ArmadaPack.bat

echo copying TPGM Templates
mkdir .\artifacts\Armada\TPGM_TEMPLATES
xcopy .\TPGM_TEMPLATES .\artifacts\Armada\TPGM_TEMPLATES /Q /E /Y

echo copying python file...
xcopy .\py_scripts .\artifacts\Armada\py_env /Q /E /Y
echo =======================================

echo Building New Project...
dotnet publish .\src\NewProject -r win-x64 -c Release --self-contained true -p:PublishSingleFile=true -o .\artifacts\Armada
echo =======================================