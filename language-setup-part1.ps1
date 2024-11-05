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
    }
    else {
        $osName
    }
    $gitlab_root = "https://gitlab.com/ondemand-engineering"
    $languagePackUri = "$gitlab_root/windows/windows-language-setup/-/raw/main/language_packs/$os/Microsoft-Windows-Server-Language-Pack_x64_$($primaryLanguage.tolower()).cab"
}

process {
    # Windows Server Language Pack Install
    if ($osName -like "*Microsoft Windows Server*") {
        foreach ($lang in ($languages | Where-Object { $_ -ne 'en-US' })) {

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

            # Remove Languae Pack
            $languagePack | Remove-Item -Force

            # Install Windows Capability
            while ($null -ne ($capabilities = Get-WindowsCapability -Online | Where-Object { $_.Name -match "$lang" -and $_.State -ne "Installed" })) {
                foreach ($capability in $capabilities) {
                    try {
                        Add-WindowsCapability -Online -Name $capability.Name
                        Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Installed $($capability.Name)" -Severity Information -LogPath $logPath
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        if ($Null -eq $errorMessage) {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to install $($capability.Name)" -Severity Error -LogPath $logPath
                        }
                        else {
                            Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                        }
                    }
                }
            }
        }
    }
    # Windows Client OS Language Install
    else {
        # Install language packs for non US
        foreach ($lang in ($languages | Where-Object { $_ -ne 'en-US' })) {
            try {
                Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Installing Language Pack" -Severity Information -LogPath $logPath
                Install-Language -Language $lang
            }
            catch {
                $errorMessage = $_.Exception.Message
                if ($Null -eq $errorMessage) {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): Failed to install Language Pack: $_" -Severity Error -LogPath $logPath
                }
                else {
                    Write-Log -Object "LanguageSetup_Part1" -Message "$($lang): $errorMessage" -Severity Error -LogPath $logPath
                }
            }
        }
    }

    # Set System Language
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

end {
    # Restart Computer
    Restart-Computer -Force
}
