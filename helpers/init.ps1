function Set-HuduInstance {
    $HuduBaseURL = $HuduBaseURL ?? 
        $((Read-Host -Prompt 'Set the base domain of your Hudu instance (e.g https://myinstance.huducloud.com)') -replace '[\\/]+$', '') -replace '^(?!https://)', 'https://'
    $HuduAPIKey = $HuduAPIKey ?? "$(read-host "Please Enter Hudu API Key")"
    while ($HuduAPIKey.Length -ne 24) {
        $HuduAPIKey = (Read-Host -Prompt "Get a Hudu API Key from $($settings.HuduBaseDomain)/admin/api_keys").Trim()
        if ($HuduAPIKey.Length -ne 24) {
            Write-Host "This doesn't seem to be a valid Hudu API key. It is $($HuduAPIKey.Length) characters long, but should be 24." -ForegroundColor Red
        }
    }
    New-HuduAPIKey $HuduAPIKey
    New-HuduBaseURL $HuduBaseURL
}
function Get-HuduModule {
    param (
        [string]$HAPImodulePath = "C:\Users\$env:USERNAME\Documents\GitHub\HuduAPI\HuduAPI\HuduAPI.psm1",
        [bool]$use_hudu_fork = $true
        )

    if ($true -eq $use_hudu_fork) {
        if (-not $(Test-Path $HAPImodulePath)) {
            $dst = Split-Path -Path (Split-Path -Path $HAPImodulePath -Parent) -Parent
            Write-Host "Using Lastest Master Branch of Hudu Fork for HuduAPI"
            $zip = "$env:TEMP\huduapi.zip"
            Invoke-WebRequest -Uri "https://github.com/Hudu-Technologies-Inc/HuduAPI/archive/refs/heads/master.zip" -OutFile $zip
            Expand-Archive -Path $zip -DestinationPath $env:TEMP -Force 
            $extracted = Join-Path $env:TEMP "HuduAPI-master" 
            if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
            Move-Item -Path $extracted -Destination $dst 
            Remove-Item $zip -Force
        }
    } else {
        Write-Host "Assuming PSGallery Module if not already locally cloned at $HAPImodulePath"
    }

    if (Test-Path $HAPImodulePath) {
        Import-Module $HAPImodulePath -Force
        Write-Host "Module imported from $HAPImodulePath"
    } elseif ((Get-Module -ListAvailable -Name HuduAPI).Version -ge [version]'2.4.4') {
        Import-Module HuduAPI
        Write-Host "Module 'HuduAPI' imported from global/module path"
    } else {
        Install-Module HuduAPI -MinimumVersion 2.4.5 -Scope CurrentUser -Force
        Import-Module HuduAPI
        Write-Host "Installed and imported HuduAPI from PSGallery"
    }
}
function Get-HuduVersionCompatible {
    param (
        [version]$RequiredHuduVersion = [version]"2.39.4",
        $DisallowedVersions = @([version]"2.37.0")
    )
    Write-Host "Required Hudu version: $requiredversion" -ForegroundColor Blue
    try {
        $HuduAppInfo = Get-HuduAppInfo
        $CurrentHuduVersion = $HuduAppInfo.version

        if ([version]$CurrentHuduVersion -lt [version]$RequiredHuduVersion) {
            Write-Host "This script requires at least version $RequiredHuduVersion and cannot run with version $CurrentHuduVersion. Please update your version of Hudu." -ForegroundColor Red
            exit 1
        }
    } catch {
        write-host "error encountered when checking hudu version for $(Get-HuduBaseURL) - $_"
    }
    Write-Host "Hudu Version $CurrentHuduVersion is compatible"  -ForegroundColor Green
}

function Get-PSVersionCompatible {
    param (
        [version]$RequiredPSversion = [version]"7.5.1"
    )

    $currentPSVersion = (Get-Host).Version
    Write-Host "Required PowerShell version: $RequiredPSversion" -ForegroundColor Blue

    if ($currentPSVersion -lt $RequiredPSversion) {
        Write-Host "PowerShell $RequiredPSversion or higher is required. You have $currentPSVersion." -ForegroundColor Red
        exit 1
    } else {
        Write-Host "PowerShell version $currentPSVersion is compatible." -ForegroundColor Green
    }
}
