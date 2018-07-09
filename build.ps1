try
{
    if (@(Get-module 'Invoke-MSBuild' -ListAvailable).count -eq 0)
    {
        install-module 'Invoke-MSBuild' -scope currentuser -confirm:$false
    }
    Invoke-MSBuild -path (Join-Path $PSScriptRoot 'SSMSThemeEditor.sln')
}
catch
{
    throw
    Read-Host "Any key to close"
}