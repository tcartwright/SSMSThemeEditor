@rem bat file to ease use of the powershell script

@%~d0
@cd "%~dp0"

powershell.exe -ExecutionPolicy RemoteSigned -NoLogo -NonInteractive -NoProfile -file "%~dp0install.ps1" 

@pause
