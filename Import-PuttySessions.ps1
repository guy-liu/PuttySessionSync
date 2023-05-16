param (
[string]$ImportFilePath = "PuttySessions.zip",
[bool]$OverwriteDuplicates = $false,
[switch]$ListOnly,
[string[]]$ExcludedSessions = @(),
[string]$tempFolderPath = "$env:TEMP\PuttySessionsExport"
 )

[string]$regSessionsPath = "HKCU:\Software\SimonTatham\PuTTY\Sessions"

if (!(Test-Path $ImportFilePath))
{
    Write-Host "Invalid import file specified. Exiting."
    Exit
}

if( Test-Path $tempFolderPath )
{
    Write-Host "Temp folder already exists. Exiting."
    Exit
}
else
{
    $null = New-Item -ItemType directory -Path $tempFolderPath
}

if (!(Test-Path $regSessionsPath) )
{
    New-Item -Force $regSessionsPath
}

# Unpack the sessions file
Expand-Archive -Path $ImportFilePath -DestinationPath $tempFolderPath


[System.Collections.ArrayList]$excludedSessionsArrayList = @()

foreach ($excludedSession in $ExcludedSessions)
{
    $excludedSessionsArrayList.Add($excludedSession.ToLower())
}


[System.Collections.ArrayList]$existingSessionsArrayList = @()

$sessions = Get-ChildItem $regSessionsPath

foreach ($session in $sessions)
{
    $sessionName = $session.PSChildName
    $null = $existingSessionsArrayList.Add($sessionName.ToLower())
}



if ($ListOnly)
{
    $regFiles = Get-ChildItem $tempFolderPath -File

    foreach ($regFile in $regFiles)
    {
        Write-Host $regFile.Name
    }
}
else
{
    $regFiles = Get-ChildItem $tempFolderPath -File

    foreach ($regFile in $regFiles)
    {
        $sessionName = $regFile.BaseName
        $sessionNameLowerCase = $sessionName.ToLower()

        if ($excludedSessionsArrayList.Contains($sessionNameLowerCase))
        {
            Write-Host "Excluding session: $sessionName"
        }
        else
        {
            if( $existingSessionsArrayList.Contains($sessionNameLowerCase))
            {
                if($OverwriteDuplicates)
                {
                    Write-Host "Overwriting duplicate session: $sessionName"
                    
                    $duplicateRegPath = "$regSessionsPath\$sessionName"
                    
                    Remove-Item $duplicateRegPath -WhatIf


                }
                else
                {
                    Write-Host "Skipping duplicate session: $sessionName"
                }
            }
            else
            {
                Write-Host "Importing session: $sessionName"

                $regFilePath = $regFile.FullName

                REG import $regFilePath
            }
        }
    }
}

Remove-Item $tempFolderPath -Recurse