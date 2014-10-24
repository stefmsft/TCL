cd /d "C:\Scratch\Projets\Tracking\TCL" &msbuild "TCL.csproj" /t:sdvViewer /p:configuration="Debug" /p:platform=Any CPU
exit %errorlevel% 