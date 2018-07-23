@rem bat file to ease use of the powershell script

@%~d0
@cd "%~dp0"

powershell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -file "%~dp0install.ps1" 

