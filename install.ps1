Set-StrictMode -Version 1
Clear-Host

function NewFolderIfNotExists($folder){
    if(!(Test-Path -Path $folder -PathType Container)) {
        New-Item -ItemType directory -Path $folder | Out-Null
    }
}

function BuildProject($dotNetProject) {
    try
    {
        if (@(Get-module 'Invoke-MSBuild' -ListAvailable).count -eq 0)
        {
            install-module 'Invoke-MSBuild' -scope currentuser -confirm:$false
        }
        Write-Host "Building VSColorThemes solution"
        $result = Invoke-MSBuild -Path $dotNetProject.FullName -MsBuildParameters "/target:Clean;Build /property:Configuration=Release" -ShowBuildOutputInNewWindow
        if (!($result.BuildSucceeded)) {
            Write-Error $result.Message
            exit
        }
    }
    catch
    {
        throw $result.Message
        if ($host.Name -match 'consolehost') {
            Read-Host -Prompt “Press Enter to exit” 
        }        
        exit
    }
}

function DownLoadThemeFiles($OutputFolder){
    $RepositoryZipUrl = "https://api.github.com/repos/Microsoft/VS-ColorThemes/zipball/master" 
    $zipFile = "$([System.IO.Path]::GetTempFileName()).zip"
 
    # download the zip 
    Write-Host "Downloading VS-ColorThemes repo files as they are not found"
    #bump tls from 1.0 to 1.2 as 1.0 is disabled now
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile 
     
    $zipDir = $zipFile -replace "\.tmp", "_tmp" -replace "\.zip", "_zip"

    #extract the zip out to the zip directory
    [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null 
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $zipDir) 

    $path = Get-ChildItem -Path "$zipDir\*\" -Directory
    Copy-Item -Path "$path\*" -Destination $OutputFolder -Force -Recurse | Out-Null

    #try to cleanup the temp files/directories 
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
        Write-Host "Restarting the powershell script as admin using: '$($script:MyInvocation.MyCommand.Path)'"
        $process = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
        $process.Arguments = "& '$($script:MyInvocation.MyCommand.Path)'";
        # Indicate that the process should be elevated
        $process.Verb = "runas";
        [System.Diagnostics.Process]::Start($process);
        Exit
    }
}

function GetThemeFiles([string]$Path){
    return Get-ChildItem -Path $Path -Filter *.* -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*.csproj" }
}

#lets start the process to deploy
Test-Admin

$dir = $PSScriptRoot
Set-Location -Path $dir

#as close as could be figured out how to dynamically query where the locations of SSMS are. TY to SHull for help
$ssmsFiles = (Get-ChildItem 'HKLM:\SOFTWARE\Classes\' `
            | Where-Object { $_.Name -match 'ssms\.sql\.\d+\.\d+' } `
            | ForEach-Object { 
                [System.IO.FileInfo]((Get-ItemProperty "$($_.PSPath)\Shell\Open\Command").'(default)' -replace " /dde", "").TrimStart('"').TrimEnd('"') 
            } `
            | Where-Object { Test-Path $_.FullName }) 

if ($ssmsFiles.Length -ge 0) {
    #grab a list of local pkgdef files. will be empty if they downloaded the zip.
    $themeFiles = GetThemeFiles -Path "$dir\VS-ColorThemes\VSColorThemes" 
    if ($themeFiles.Length -eq 0) {
        #down load the zip of VSColorThemes repo and build it so we can grab their pkgdef files
        DownLoadThemeFiles -OutputFolder "$dir\VS-ColorThemes" | Where-Object { $_.Name -notlike "*.csproj" }
        $themeFiles = GetThemeFiles -Path "$dir\VS-ColorThemes\VSColorThemes" 
    }

    #grab the pkgdefs from the build except for the ones I override
    $pkgDefFiles = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes\bin\Release\Themes" -Filter *.pkgdef -File -ErrorAction SilentlyContinue
    $themeXmlFiles = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes\Themes" -Filter *.xml -File -ErrorAction SilentlyContinue

    #if we dont have the same number of pkgdef files that we do xml files we need to build
    if ($pkgDefFiles.Length -ne $themeXmlFiles.Length) {
        #we must build the project to create the pkgdef files if they do not exist
        $proj = Get-ChildItem -Path "$dir\VS-ColorThemes\VSColorThemes\*.csproj" -File 
        BuildProject -dotNetProject $proj
    }
    #grab the override package defs
    $pkgDefFilesTemp = Get-ChildItem -Path "$dir\SSMSThemeEditor\SSMSPkgDefThemes" -Filter *.pkgdef -File
    #remove all the package defs that are in the overrides
    $pkgDefFiles = $pkgDefFiles | Where-Object { 
        $f = $_
        -not ($pkgDefFilesTemp | Where-Object { $f.Name -like $_.Name })  
    }
    #add the custom override pkgdef themes into the main array
    $pkgDefFiles += $pkgDefFilesTemp;

    $themeSubFolder = "VSColorThemes\"

    Write-Host "Installing themes for all installed SSMS applications:`r`n"
    foreach ($ssmsFile in $ssmsFiles) {
        $themeDir = "$($ssmsFile.DirectoryName)\Extensions\$themeSubFolder"
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