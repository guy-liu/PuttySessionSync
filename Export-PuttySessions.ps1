param (
[string]$exportFilePath = "PuttySessions.zip",
[string]$tempFolderPath = "$env:TEMP\PuttySessionsExport",
[string]$regSessionsPath = "HKCU:\Software\SimonTatham\PuTTY\Sessions",
[string]$RegistryFilePath = "" # i.e. "NTUser.dat"
 )

if (!(Test-Path $regSessionsPath) )
{
    Write-Host "Did not find any saved sessions. Exiting."
    Exit
}

if (Test-Path $exportFilePath)
{
    Write-Host "Destination file already exists. Exiting."
    Exit
}

if( Test-Path $tempFolderPath )
{
    Write-Host "Temp folder already exists. Exiting."
    Exit
}
else
{
    New-Item -ItemType directory -Path $tempFolderPath
}

# Mount registry file if specified
if($RegistryFilePath -ne "")
{
    
    if (!(Test-Path $RegistryFilePath))
    {
        Write-Host "Invalid registry file specified. Exiting."
        Exit
    }


    REG LOAD HKU\ExportPuttySessions $RegistryFilePath

    $regSessionsPath = "HKU:\ExportPuttySessions\Software\SimonTatham\PuTTY\Sessions"

    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
}


# Go through all the sessions
$sessions = Get-ChildItem $regSessionsPath

foreach ($session in $sessions)
{
    
    $sessionName = $session.PSChildName
    $sessionRegPath = $session.Name
    $tempRegExportPath = "$tempFolderPath\$sessionName.reg"
    REG EXPORT $sessionRegPath $tempRegExportPath

    # Fix Reg key path if loaded from hive
    if ( $RegistryFilePath -ne $null)
    {
        $content = Get-Content $tempRegExportPath

        $content = $content -creplace "HKEY_USERS\\ExportPuttySessions", "HKEY_CURRENT_USER"

        $content | Set-Content $tempRegExportPath
    }

    $session.Close()
}

Compress-Archive -Path "$tempFolderPath\*" -DestinationPath $exportFilePath

Remove-Item $tempFolderPath -Recurse


# Clean up if registry file is mounted
if($RegistryFilePath -ne "")
{
    $sessions = $null

    Remove-PSDrive HKU
    
    [GC]::Collect()

    REG UNLOAD  HKU\ExportPuttySessions
}