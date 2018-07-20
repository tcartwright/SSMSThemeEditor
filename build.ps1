try
{
    if (@(Get-module 'Invoke-MSBuild' -ListAvailable).count -eq 0)
    {
        install-module 'Invoke-MSBuild' -scope currentuser -confirm:$false
    }
    
    $slnPath = [io.path]::combine($PSScriptRoot, '..', 'SSMSThemeEditor.sln')

    Invoke-MSBuild -path $slnPath
}
catch
{
    throw
    Read-Host "Any key to close"
}