Set-StrictMode -Version 1
Clear-Host

## self elevating code source: https://blogs.msdn.microsoft.com/virtual_pc_guy/2010/09/23/a-self-elevating-powershell-script/
# Get the ID and security principal of the current user account
$winID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$winPrincipal=new-object System.Security.Principal.WindowsPrincipal($winID)
 
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
# Check to see if we are currently running "as Administrator"
if (!($winPrincipal.IsInRole($adminRole))) {
   $process = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   $process.Arguments = $myInvocation.MyCommand.Definition;
   # Indicate that the process should be elevated
   $process.Verb = "runas";
   [System.Diagnostics.Process]::Start($process);
   Exit
}


$dir = $PSScriptRoot
Set-Location -Path $dir

#would really prefer to query installed locations of ssms but it is not stored in registry as far as I can find
Write-Host "Scanning $Env:ProgramFiles\ for ssms.exe"
$files = Get-ChildItem -Path "$Env:ProgramFiles\" -Filter "ssms.exe" -Recurse
Write-Host "Scanning ${Env:ProgramFiles(x86)}\ for ssms.exe"
$files += Get-ChildItem -Path "${Env:ProgramFiles(x86)}\" -Filter "ssms.exe"  -Recurse

if ($files.Length -ge 0) {
    $themeFiles = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes" -Filter *.* -File | Where-Object { $_.Name -notlike "*.csproj" }
    $pkgDefFiles = Get-ChildItem -Path "$dir\SSMSThemeEditor\SSMSPkgDefThemes" -Filter *.* -File
    $themeSubFolder = "VSColorThemes\"

    Write-Host "Installing themes for all installed SSMS applications:`r`n"
    foreach ($file in $files) {
        $themeDir = "$($file.Directory)\Extensions\$themeSubFolder"
        Write-Host "Installing theme extension to $($themeDir)";

        if(!(Test-Path -Path $themeDir -PathType Container)) {
            New-Item -ItemType directory -Path $themeDir | Out-Null
        }
        foreach($themeFile in $themeFiles) {
            Write-Host "Deploying $($themeSubFolder)$($themeFile.Name)"
            Copy-Item -Path $themeFile.FullName -Destination $themeDir -Force | Out-Null
        }

        $themesDir = "$($themeDir)Themes\"
        if(!(Test-Path -Path $themesDir -PathType Container)) {
            New-Item -ItemType directory -Path $themesDir | Out-Null
        }
        foreach($pkgDefFile in $pkgDefFiles) {
            Write-Host "Deploying $($themeSubFolder)Themes\$($pkgDefFile.Name)"
            Copy-Item -Path $pkgDefFile.FullName -Destination $themesDir -Force | Out-Null
        }

        Write-Host "Done`r`n"
    }
} else {
    Write-Host "No installations of SSMS found!"
}

Write-Host "Complete`r`n"

Read-Host -Prompt “Press Enter to exit” 