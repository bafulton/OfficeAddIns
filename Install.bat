@echo off

echo Installing:

copy /y /v ".\Word\Normal.dotm" "%HOMEDRIVE%%HOMEPATH%\AppData\Roaming\Microsoft\Templates\Normal.dotm"
copy /y /v ".\Excel\bin\Blackman and Sloop Add-In v1.4.2.xlam" "%HOMEDRIVE%%HOMEPATH%\AppData\Roaming\Microsoft\AddIns\Blackman and Sloop Add-In.xlam"

echo:
echo Installation success!
echo:
pause

@echo on