@echo off

set iDir=.\src\ArmadaPack
set oDir=.\artifacts\Armada

echo ========== Building x64 ArmadaPack =======================

dotnet build %iDir% --arch x64 --output %oDir%

echo ========== Building x32 ArmadaPack =======================

set oDir=.\TPGM_TEMPLATES\TPGM_Golden
dotnet build %iDir% --arch x86 --output %oDir%

set oDir=.\TPGM_TEMPLATES\TPGM_Test
dotnet build %iDir% --arch x86 --output %oDir%