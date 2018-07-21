Set-StrictMode -Version 1
Clear-Host

function NewFolderIfNotExists($folder){
    if(!(Test-Path -Path $folder -PathType Container)) {
        New-Item -ItemType directory -Path $folder | Out-Null
    }
}

function DownLoadThemeFiles($OutputFolder){
    #bump tls from 1.0 to 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $RepositoryZipUrl = "https://api.github.com/repos/Microsoft/VS-ColorThemes/zipball/master" 
    $zipFile = "$([System.IO.Path]::GetTempFileName()).zip"
 
    # download the zip 
    Write-Host "Downloading VS-ColorThemes repo files as they are not found"
    Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile 
     
    $zipDir = "$([System.IO.Path]::GetDirectoryName($zipFile))\$([System.IO.Path]::GetFileNameWithoutExtension($zipFile))" -replace ".tmp", "_tmp"
    NewFolderIfNotExists -folder $zipDir
    NewFolderIfNotExists -folder $OutputFolder

    [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null 
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $zipDir) 

    $zipFiles = Get-ChildItem -Path "$zipDir\*\VSColorThemes\*.*" -File


    foreach($zipFile in $zipFiles) {
        Copy-Item -Path $zipFile.FullName -Destination $OutputFolder -Force | Out-Null
    }

    Remove-Item -Path $zipFile -Force -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path $zipDir -Force -Recurse -ErrorAction SilentlyContinue | Out-Null
}

function Test-Admin() {
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
}

Test-Admin

$dir = $PSScriptRoot
Set-Location -Path $dir

#would really prefer to query installed locations of ssms but it is not stored in registry as far as I can find
Write-Host "Scanning $Env:ProgramFiles\ for ssms.exe"
$files = Get-ChildItem -Path "$Env:ProgramFiles\" -Filter "ssms.exe" -Recurse
Write-Host "Scanning ${Env:ProgramFiles(x86)}\ for ssms.exe"
$files += Get-ChildItem -Path "${Env:ProgramFiles(x86)}\" -Filter "ssms.exe"  -Recurse

if ($files.Length -ge 0) {
    $themeFiles = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes" -Filter *.* -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*.csproj" }
    if ($themeFiles.Length -eq 0) {
        DownLoadThemeFiles -OutputFolder "$dir\VS-ColorThemes\VSColorThemes" | Where-Object { $_.Name -notlike "*.csproj" }
        $themeFiles = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes" -Filter *.* -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*.csproj" }
    }

    $pkgDefFiles = Get-ChildItem -Path "$dir\SSMSThemeEditor\SSMSPkgDefThemes" -Filter *.* -File
    $themeSubFolder = "VSColorThemes\"

    Write-Host "Installing themes for all installed SSMS applications:`r`n"
    foreach ($file in $files) {
        $themeDir = "$($file.Directory)\Extensions\$themeSubFolder"
        Write-Host "Installing theme extension to $($themeDir)";

        NewFolderIfNotExists -folder $themeDir
        foreach($themeFile in $themeFiles) {
            Write-Host "Deploying $($themeSubFolder)$($themeFile.Name)"
            Copy-Item -Path $themeFile.FullName -Destination $themeDir -Force | Out-Null
        }

        $themesDir = "$($themeDir)Themes\"
        NewFolderIfNotExists -folder $themesDir
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

if ($host.Name -match 'consolehost') {
    Read-Host -Prompt “Press Enter to exit” 
}