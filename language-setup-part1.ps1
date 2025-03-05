<#
    .DESCRIPTION
    Language Setup Part 1
#>

Param (
    [Parameter(Mandatory = $false, HelpMessage = 'Primary Language')]
    [ValidateNotNullorEmpty()]
    [string] $primaryLanguage = "en-GB",

    [Parameter(Mandatory = $false, HelpMessage = 'Secondary Language')]
    [ValidateNotNullOrEmpty()]
    [String] $secondaryLanguage = "en-US",

    [Parameter(Mandatory = $false, HelpMessage = 'Additional Language')]
    [ValidateNotNullOrEmpty()]
    [string] $additionalLanguages
)

begin {
    function Write-Log {
        [CmdletBinding()]
        <#
            .SYNOPSIS
            Create log function
        #>
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $logPath,

            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $object,

            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [System.String] $message,

            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('Information', 'Warning', 'Error', 'Verbose', 'Debug')]
            [System.String] $severity,

            [Parameter(Mandatory = $False)]
            [Switch] $toHost
        )

        begin {
            $date = (Get-Date).ToLongTimeString()
        }
        process {
            if (($severity -eq "Information") -or ($severity -eq "Warning") -or ($severity -eq "Error") -or ($severity -eq "Verbose" -and $VerbosePreference -ne "SilentlyContinue") -or ($severity -eq "Debug" -and $DebugPreference -ne "SilentlyContinue")) {
                if ($True -eq $toHost) {
                    Write-Host $date -ForegroundColor Cyan -NoNewline
                    Write-Host " - [" -ForegroundColor White -NoNewline
                    Write-Host "$object" -ForegroundColor Yellow -NoNewline
                    Write-Host "] " -ForegroundColor White -NoNewline
                    Write-Host ":: " -ForegroundColor White -NoNewline

                    Switch ($severity) {
                        'Information' {
                            Write-Host "$message" -ForegroundColor White
                        }
                        'Warning' {
                            Write-Warning "$message"
                        }
                        'Error' {
                            Write-Host "ERROR: $message" -ForegroundColor Red
                        }
                        'Verbose' {
                            Write-Verbose "$message"
                        }
                        'Debug' {
                            Write-Debug "$message"
                        }
                    }
                }
            }

            switch ($severity) {
                "Information" { [int]$type = 1 }
                "Warning" { [int]$type = 2 }
                "Error" { [int]$type = 3 }
                'Verbose' { [int]$type = 2 }
                'Debug' { [int]$type = 2 }
            }

            if (!(Test-Path (Split-Path $logPath -Parent))) { New-Item -Path (Split-Path $logPath -Parent) -ItemType Directory -Force | Out-Null }

            $content = "<![LOG[$message]LOG]!>" + `
                "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " + `
                "date=`"$(Get-Date -Format "M-d-yyyy")`" " + `
                "component=`"$object`" " + `
                "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + `
                "type=`"$type`" " + `
                "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " + `
                "file=`"`">"
            if (($severity -eq "Information") -or ($severity -eq "Warning") -or ($severity -eq "Error") -or ($severity -eq "Verbose" -and $VerbosePreference -ne "SilentlyContinue") -or ($severity -eq "Debug" -and $DebugPreference -ne "SilentlyContinue")) {
                Add-Content -Path $($logPath + ".log") -Value $content
            }
        }
        end {}
    }

    $logPath = "$env:SYSTEMROOT\TEMP\Deployment_" + (Get-Date -Format 'yyyy-MM-dd')

    [array]$languages = $primaryLanguage, $secondaryLanguage
    if (!([string]::IsNullOrEmpty($additionalLanguages))) {
        $languages += $additionalLanguages.Split(';')
    }

    # Get OS Name
    $osName = (Get-ComputerInfo).OsName
    $os = if ($osName -match "Server \d+") {
        $matches[0].Replace(" ", "_").tolower()
        $type = "Server"
    }
    elseif ($osName -match "Windows \d+") {
        $matches[0].Replace(" ", "_").tolower()
        $type = "Client"
    }
    else {
        $osName
    }
    $storage_account = "https://mcduksstoracc001.blob.core.windows.net"
    $blob_root = "$storage_account/media/windows/language_packs/$os"

    $reboot = $false
}

process {
    # Disable Language Pack Cleanup
    Disable-ScheduledTask -TaskPath "\Microsoft\Windows\AppxDeploymentClient\" -TaskName "Pre-staged app cleanup"
    Disable-ScheduledTask -TaskPath "\Microsoft\Windows\MUI\" -TaskName "LPRemove"
    Disable-ScheduledTask -TaskPath "\Microsoft\Windows\LanguageComponentsInstaller" -TaskName "Uninstallation"
    reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Control Panel\International" /v "BlockCleanupOfUnusedPreinstalledLangPacks" /t REG_DWORD /d 1 /f

    foreach ($lang in ($languages | Where-Object { $_ -ne 'en-US' })) {

        if (!(Get-WindowsPackage -Online | Where-Object { $_.ReleaseType -eq "LanguagePack" -and $_.PackageName -like "*LanguagePack*$lang*" })) {

            $languagePackUri = "$blob_root/Microsoft-Windows-$type-Language-Pack_x64_$($lang.toLower()).cab"

            # Download Language Pack
            try {
                Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Downloading Language Pack" -Severity Information -LogPath $logPath
                Start-BitsTransfer -Source $languagePackUri -Destination "$env:SYSTEMROOT\Temp\$(Split-Path $languagePackUri -Leaf)"
                Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Downloaded Language Pack" -Severity Information -LogPath $logPath
                $languagePack = Get-Item -Path "$env:SYSTEMROOT\Temp\$(Split-Path $languagePackUri -Leaf)"
                Unblock-File -Path $languagePack.FullName -ErrorAction SilentlyContinue
            }
            catch {
                $errorMessage = $_.Exception.Message
                if ($Null -eq $errorMessage) {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to Download Language Pack: $_" -Severity Error -LogPath $logPath
                }
                else {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                }
            }

            # Install Language Pack
            Try {
                Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Installing Language Pack" -Severity Information -LogPath $logPath
                Add-WindowsPackage -Online -PackagePath $languagePack.FullName -NoRestart
                Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Installed Language Pack" -Severity Information -LogPath $logPath
                $reboot = $true
            }
            Catch {
                $errorMessage = $_.Exception.Message
                if ($Null -eq $errorMessage) {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to install Language Pack: $_" -Severity Error -LogPath $logPath
                }
                else {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                }
            }

            # Remove Language Pack file
            $languagePack | Remove-Item -Force
        }

        if (($os -ne "server_2016") -or ($os -ne "server_2019")) {
            $capabilities = @(
                "Microsoft-Windows-LanguageFeatures-Basic-$($lang.toLower())-Package~31bf3856ad364e35~amd64~~.cab",
                "Microsoft-Windows-LanguageFeatures-Handwriting-$($lang.toLower())-Package~31bf3856ad364e35~amd64~~.cab",
                "Microsoft-Windows-LanguageFeatures-OCR-$($lang.toLower())-Package~31bf3856ad364e35~amd64~~.cab",
                "Microsoft-Windows-LanguageFeatures-Speech-$($lang.toLower())-Package~31bf3856ad364e35~amd64~~.cab",
                "Microsoft-Windows-LanguageFeatures-TextToSpeech-$($lang.toLower())-Package~31bf3856ad364e35~amd64~~.cab"
            )

            foreach ($capability in $capabilities) {

                if ((Get-WindowsCapability -Online | Where-Object { $_.Name -match "$lang" -and $_.Name -match $capability.Split("-")[3] }).State -ne "Installed") {

                    $capabilityUri = "$blob_root/$capability"

                    # Download Windows Capability
                    try {
                        Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Downloading $capability" -Severity Information -LogPath $logPath
                        Start-BitsTransfer -Source $capabilityUri -Destination "$env:SYSTEMROOT\Temp\$(Split-Path $capabilityUri -Leaf)"
                        Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Downloaded $capability" -Severity Information -LogPath $logPath
                        $file = Get-Item -Path "$env:SYSTEMROOT\Temp\$(Split-Path $capabilityUri -Leaf)"
                        Unblock-File -Path $file.FullName -ErrorAction SilentlyContinue
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        if ($Null -eq $errorMessage) {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to Download Language Pack: $_" -Severity Error -LogPath $logPath
                        }
                        else {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                        }
                    }

                    # Install Windows Capability
                    try {
                        Add-WindowsPackage -Online -PackagePath $file.FullName -NoRestart
                        Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Installed $capability" -Severity Information -LogPath $logPath
                        $reboot = $true
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        if ($Null -eq $errorMessage) {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to install $capability" -Severity Error -LogPath $logPath
                        }
                        else {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                        }
                    }

                    # Remove Windows Capability file
                    $file | Remove-Item -Force
                }
            }
        }
    }

    # Set System Language
    if ((Get-WinSystemLocale).Name -ne $primaryLanguage) {
        try {
            Set-WinSystemLocale -SystemLocale $primaryLanguage
            Write-Log -Object "LanguageSetup_Part1" -Message "Set System Locale to $primaryLanguage" -Severity Information -LogPath $logPath
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($Null -eq $errorMessage) {
                Write-Log -Object "LanguageSetup_Part1" -Message "Failed to set System Locale to $primaryLanguage" -Severity Error -LogPath $logPath
            }
            else {
                Write-Log -Object "LanguageSetup_Part1" -Message "$errorMessage" -Severity Error -LogPath $logPath
            }
        }
    }
}

end {
    # Restart Computer
    if ($reboot) {
        Restart-Computer -Force
    }
}
