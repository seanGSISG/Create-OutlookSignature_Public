<#
.SYNOPSIS  
Generates an HTML email signature for Microsoft Outlook using Base64 encoded images using the Microsoft Graph API to collect user information from Azure AD.
Must register an application in Azure AD and grant the necessary permissions to read user information.

.NOTES  
Author: SeanGSISG
Version: 1.0.0
Last Updated: December 20, 2024
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$User,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Default', 'SecondaryDomain', 'Alternative')]
    [string]$Template = 'Default',
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Logo-Default', 'Logo-SecondaryDomain', 'Logo-Alternative', 'Logo-Default2', 'Logo-Default3')]
    [string]$Logo = "Logo-Default",
    
    [switch]$Cleanup
)

# Version check
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "This script requires PowerShell 5 or higher."
    exit 1
}

Write-Host "Running in PowerShell $($PSVersionTable.PSVersion)"

#region Configuration
# Initialize script variables
$ErrorActionPreference = "Stop"
$logBasePath = "C:\Temp\Signatures"
if (-not (Test-Path $logBasePath)) {
    New-Item -ItemType Directory -Path $logBasePath -Force | Out-Null
    Write-Host "Created log directory: $logBasePath"
}
$logFilePath = Join-Path -Path $logBasePath -ChildPath "OutlookSignature_$(Get-Date -Format 'yyyyMMdd').log"

# Graph API Configuration
$authConfig = @{
    TenantId = "Your-Tenant-ID"
    ClientId = "Your-Client-ID"
    ClientSecret = "Your-Secret"
    Scope = "https://graph.microsoft.com/.default"
}

# Company Addresses
$Addresses = @{
    "Office_1"      = "123 Main Street, Honolulu, HI 96813"
    "Office_2"      = "123 Main Blvd, Boulder, CO 80302"
    "Office_3"      = "123 Main Court, Zanzabar, Guam 65482"
    "DEFAULT"       = "123 Main Your, Moms, HOUSE 696969"
}

# Base64 Logos
$Base64Logos = @{
    "Logo-Default"          = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJYAAABECAYAAABj98zGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKtSURBVHhe7dbNjRQxFMTxiYIzCRAEIgiunIkBMiAAzmRAGGRCFoPeobRPr9y77FDqD/Q//KTussfjtt+053b78fsOxFkAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFgAJFiAc/n8/X673e63D5+87cws2FstWpn5Xt69f5qDvHnr/Y5CYT3o6MKqIqrv//bL286AwnrQc4VVi9nfJB+/PrV9+elvmj7OzLcKZ6uwNH690fq9NlgbPudW4+heY5fqX9nsr3GrvfevvH9PL6y5Lup7JhbsTYszcy2eFk2bUgutz2mx1VfFUdf9ONOGze/obbOwio7JmoOuq1/NQQXX+1WbCqvPXfcqpt6//0D6s2m+s7DmumisOfejWbC3voiizemb17P+K682LX7lswDLKpP+lhBtWi+S0t+YXd9sfWZV2Cre3n++GWf7LKw5ttpXz3YkC/amTevZS4Wl6/krrutVEa0ymZs+9T/3vc9WQb6msGo+zxVWta8Ka2X1bEeyYG9amJ69VFj9c7J1ZG5lMje968dUmYU8CyVdWH/zxjorC/amTZt5/x9S96vFVjF1swD1HVubMTd91dav6zvn3Hrbo4XVn0f3dT0Lq6/DnO+ZWLC3rSNl1dYXc7YVbdx802wVVelHXe+vDdX/Kt33o1g0xr8UVo3Rn0l9Z2FtzXk+19EsuAIdbSq01WZexeoo/B9YcAXzOLjy5lx57s+x4CrmUXjVjaGwgFewAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAEiwAAj4A452IDEGOZpqAAAAAElFTkSuQmCC"
    "Logo-Default2"         = "data:image/png;base64,iVBORw0KGgo....."  # Shortened for brevity
    "Logo-Default3"         = "data:image/png;base64,iVBORw0KGgo....."  # Shortened for brevity
    "Logo-SecondaryDomain"  = "data:image/png;base64,iVBORw0KGgo....."  # Shortened for brevity
    "Logo-Alternative"      = "data:image/png;base64,iVBORw0KGgo....."  # Shortened for brevity
}

# Domain Constants
$DOMAINS = @{
    Primary = "@PrimaryDomain.com"
    Secondary = "@SecondaryDomain.com"
}

# File name patterns
$FILE_PATTERNS = @{
    Default = @{
        New = "New ({0}).htm"
        Reply = "Reply ({0}).htm"
    }
    CustomLogo = @{
        New = "New_{0} ({1}).htm"  # Will produce "New_Logo-SecondaryDomain" when Logo is "Logo-SecondaryDomain"
        Reply = "Reply_{0} ({1}).htm"  # Will produce "Reply_SecondaryDomain" when Logo is "Logo-SecondaryDomain"
    }
}

# Template mapping
$TEMPLATE_PATHS = @{
    Default = "templates\Default"
    SecondaryDomain = "templates\SecondaryDomain"
    Alternative = "templates\Alternative"
}

function Get-SignatureFileName {
    param (
        [string]$EmailAddress,
        [string]$Type,
        [string]$Logo
    )
    
    if ([string]::IsNullOrEmpty($Logo) -or $Logo -eq "Logo-Default") {
        return [string]::Format($FILE_PATTERNS.Default.$Type, $EmailAddress)
    }
    else {
        return [string]::Format($FILE_PATTERNS.CustomLogo.$Type, $Logo, $EmailAddress)
    }
}

function Get-SignatureTemplate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Logo,
        [bool]$HasSecondaryDomainMailbox
    )
    
    # Determine template type based on conditions
    $templateType = if ($HasSecondaryDomainMailbox) {
        "SecondaryDomain"
    }
    elseif ($Logo -eq "Logo-Alternative") {
        "Alternative"
    }
    else {
        "Default"
    }

    # Return template paths with .htm extension
    return @{
        Path = $TEMPLATE_PATHS[$templateType]
        Templates = @{
            New = Join-Path $PSScriptRoot "$($TEMPLATE_PATHS[$templateType])\New.htm"
            Reply = Join-Path $PSScriptRoot "$($TEMPLATE_PATHS[$templateType])\Reply.htm"
        }
    }
}
#endregion

#region Helper Functions
# Authentication & Graph API
function Get-GraphAccessToken {
    param ($config, [switch]$ForceRefresh)
    
    if (-not $script:tokenCache -or $ForceRefresh) {
        try {
            $tokenResponse = Invoke-RestMethod `
                -Method Post `
                -Uri "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token" `
                -Body @{
                    grant_type = "client_credentials"
                    client_id = $config.ClientId
                    client_secret = $config.ClientSecret
                    scope = $config.Scope
                } `
                -ContentType "application/x-www-form-urlencoded"

            $script:tokenCache = @{
                Token = $tokenResponse.access_token
                Expiry = [DateTime]::UtcNow.AddSeconds($tokenResponse.expires_in)
            }
        }
        catch { throw "Failed to get access token: $_" }
    }
    return $script:tokenCache.Token
}

# Retrieve user information from Graph API
function Get-GraphUserInfo {
    param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    
    try {
        $headers = @{ Authorization = "Bearer $accessToken" }
        $userInfo = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName" `
            -Headers $headers

        # Transform office location to standard format
        $standardOfficeLocation = Format-OfficeLocation -officeLocation $userInfo.officeLocation

        return @{
            DisplayName = $userInfo.displayName.Trim()
            JobTitle = if ($userInfo.jobTitle -eq $null) { "No Title" } else { Format-Title $userInfo.jobTitle }
            EmailAddress = $userInfo.mail.Trim()
            TelephoneNumber = if ($userInfo.businessPhones[0] -eq $null) { "No Phone Number" } else { Format-PhoneNumber $userInfo.businessPhones[0] }
            OfficeLocation = $standardOfficeLocation
        }
    }
    catch {
        Write-Error "Failed to get user info: $($_.Exception.Message)"
        throw
    }
}

# Check and retrieve SecondaryDomain account information
function Get-SecondaryDomainUserInfo {
    param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    try {
        $username = $userPrincipalName.Split('@')[0]
        # Always use capitalized version for display and storage
        $SecondaryDomainEmailProperCase = "$username@SecondaryDomain.com"
        # Use lowercase only for Graph API query
        $SecondaryDomainEmailForQuery = "$username@secondarydomain.com"
        
        $headers = @{ Authorization = "Bearer $accessToken" }
        $response = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$SecondaryDomainEmailForQuery" `
            -Headers $headers `
            -ErrorAction SilentlyContinue
            
        if ($response) {
            Write-Host "Found SecondaryDomain account: $SecondaryDomainEmailProperCase"
            return [PSCustomObject]@{
                HasAccount = $true
                Email = $SecondaryDomainEmailProperCase  # Use properly capitalized version
            }
        }
    }
    catch {
        Write-Host "No SecondaryDomain account found for user"
        return [PSCustomObject]@{
            HasAccount = $false
            Email = $null
        }
    }
}

# Text Formatting & Processing
function Format-Title {
    param ([string]$Title)
    if ([string]::IsNullOrEmpty($Title)) { return "" }
    
    $lowercaseWords = @(
        "a", "an", "and", "as", "at", "but", "by", "for", "if", "in", 
        "of", "on", "or", "the", "to", "up", "yet", "into", "onto", "with"
    )  # Common articles, conjunctions, and prepositions always lowercase in user titles

    $formattedTitle = $Title -split ' ' | ForEach-Object {
        if ([string]::IsNullOrEmpty($_)) { return "" }
        if ($_ -cmatch '^(IT|IV|III|II|I|US|USA|CEO|CFO|COO|CTO|VP|PM|SME|DOD|DOE|EPA|OSHA|QA|QC)$') { 
            # Common business/government acronyms & Roman numerals always uppercase in user titles
            $_
        } else {
            if ($lowercaseWords -contains $_.ToLower()) {
                $_.ToLower()
            } else {
                if ($_.Length -gt 0) {
                    $_.Substring(0, 1).ToUpper() + $(if($_.Length -gt 1) { $_.Substring(1).ToLower() } else { "" })
                } else {
                    ""
                }
            }
        }
    }
    return ($formattedTitle -join ' ').Trim()
}

# Format text for secondary domain
function Format-SecondaryDomainText {
    param (
        [string]$text,
        [ValidateSet('Title', 'Company', 'Email', 'Address')]
        [string]$type
    )
    
    switch ($type) {
        'Title' { return $text.ToUpper() }
        'Company' { return "SECONDARYDOMAIN, INC." }
        'Email' { return $text.ToUpper() }
        'Address' { return $text.ToUpper() }
        Default { return $text }
    }
}

function Format-PhoneNumber {
    param (
        [string]$phoneNumber,
        [string]$formatType = "Primary"  # Default to primary format
    )
    
    if ([string]::IsNullOrEmpty($phoneNumber) -or $phoneNumber -eq "No Phone Number") { 
        return "No Phone Number" 
    }
    
    # Extract only digits
    $digits = $phoneNumber -replace '\D', ''
    
    if ($digits.Length -eq 10) {
        if ($formatType -eq "SecondaryDomain") {
            # Format for SecondaryDomain: XXX.XXX.XXXX
            return "$($digits.Substring(0,3)).$($digits.Substring(3,3)).$($digits.Substring(6,4))"
        } else {
            # Standard format: (XXX) XXX-XXXX
            return "($($digits.Substring(0,3))) $($digits.Substring(3,3))-$($digits.Substring(6,4))"
        }
    }
    # Return original if unable to format
    return $phoneNumber
}

# Standardize office location format
function Format-OfficeLocation {
    param ([string]$officeLocation)
    
    if ($officeLocation) {
        $location = $officeLocation.Trim().ToUpper().Replace(" ", "")
        if ($Addresses.ContainsKey($location)) {
            return $location
        }
    }
    return "DEFAULT"  # Return default if no valid location found
}

# Registry Management
function Update-OutlookRegistry {
    param (
        [string]$newEmailSignature,
        [string]$replySignature,
        [string]$userEmail
    )
    
    try {
        # Define registry paths
        $outlookVersions = @(
            "16.0"  # Outlook 2016, 2019, 2021, 365
            "15.0"  # Outlook 2013
            "14.0"  # Outlook 2010
        )

        $baseProfilePath = "Software\Microsoft\Office"
        
        foreach ($version in $outlookVersions) {
            $profilesPath = "HKCU:\$baseProfilePath\$version\Outlook\Profiles"
            if (-not (Test-Path $profilesPath)) { continue }
            
            Write-Host "Processing Outlook version $version..."
            
            # Set ShowAutoSig in Mail Options for this version
            $mailOptionsPath = "HKCU:\$baseProfilePath\$version\Outlook\Options\Mail"
            if (-not (Test-Path $mailOptionsPath)) {
                New-Item -Path $mailOptionsPath -Force | Out-Null
            }
            Set-ItemProperty -Path $mailOptionsPath -Name "ShowAutoSig" -Value 1 -Type DWord -Force
            
            # Process all profiles
            Get-ChildItem $profilesPath | ForEach-Object {
                $profileName = Split-Path $_.Name -Leaf
                $accountsPath = "$profilesPath\$profileName\9375CFF0413111d3B88A00104B2A6676"
                
                if (Test-Path $accountsPath) {
                    Get-ChildItem $accountsPath | ForEach-Object {
                        $accountPath = $_.PSPath
                        try {
                            $accountProps = Get-ItemProperty -Path $accountPath -ErrorAction SilentlyContinue
                            if ($accountProps.'Account Name' -eq $userEmail) {
                                # Set signature settings
                                Set-ItemProperty -Path $accountPath -Name "New Signature" -Value $newEmailSignature -Type String
                                Set-ItemProperty -Path $accountPath -Name "Reply-Forward Signature" -Value $replySignature -Type String
                                
                                # Additional settings for signature functionality
                                Set-ItemProperty -Path $accountPath -Name "ShowAutoSig" -Value 1 -Type DWord
                                Set-ItemProperty -Path $accountPath -Name "SignatureAdded" -Value 1 -Type DWord

                                # Enable HTML signatures
                                $setupPath = "HKCU:\$baseProfilePath\$version\Outlook\Setup"
                                if (Test-Path $setupPath) {
                                    # Enable roaming signatures
                                    Set-ItemProperty -Path $setupPath -Name "DisableRoamingSignatures" -Value 0 -Type DWord
                                    Set-ItemProperty -Path $setupPath -Name "DisableRoamingSignaturesTemporaryToggle" -Value 0 -Type DWord
                                    Set-ItemProperty -Path $setupPath -Name "DisableHTMLSignatures" -Value 0 -Type DWord
                                    Set-ItemProperty -Path $setupPath -Name "EnableRoamingSettings" -Value 1 -Type DWord
                                }

                                # New Outlook settings
                                $newOutlookPath = "HKCU:\$baseProfilePath\$version\Outlook\NewOutlook"
                                if (Test-Path $newOutlookPath) {
                                    Set-ItemProperty -Path $newOutlookPath -Name "EnableHTMLSignatures" -Value 1 -Type DWord
                                    Set-ItemProperty -Path $newOutlookPath -Name "SignatureSyncEnabled" -Value 1 -Type DWord
                                }

                                # OWA settings
                                $owaPath = "HKCU:\$baseProfilePath\$version\Common\Identity\Identities"
                                if (Test-Path $owaPath) {
                                    Get-ChildItem $owaPath | ForEach-Object {
                                        Set-ItemProperty -Path $_.PSPath -Name "EnableWebSignatures" -Value 1 -Type DWord
                                    }
                                }

                                # Signature backup settings
                                $backupPath = "HKCU:\$baseProfilePath\$version\Outlook\Options\General"
                                if (Test-Path $backupPath) {
                                    Set-ItemProperty -Path $backupPath -Name "SignatureBackupEnabled" -Value 1 -Type DWord
                                    Set-ItemProperty -Path $backupPath -Name "SignatureBackupLocation" -Value "$env:USERPROFILE\Documents\Outlook Signatures Backup" -Type String
                                }

                                Write-Host "Updated signature settings for $userEmail in Outlook $version"
                            }
                        }
                        catch {
                            Write-Warning "Error processing account registry path: $accountPath"
                            Write-Warning $_.Exception.Message
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Failed to set Outlook signature registry settings: $($_.Exception.Message)"
        throw
    }
}

# Add this to the Registry Management section, after the other registry settings
function Set-OutlookPreferences {
    try {
        $preferencesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
        
        # Create the path if it doesn't exist
        if (-not (Test-Path $preferencesPath)) {
            New-Item -Path $preferencesPath -Force | Out-Null
        }
        
        # Set UseNewOutlook to 0
        Set-ItemProperty -Path $preferencesPath -Name "UseNewOutlook" -Value 0 -Type DWord -Force
        
        # Remove specified registry values if they exist
        $valuesToRemove = @(
            "NewOutlookRenudgeStartDate",
            "NewmailDesktopAlertsDRMPreview"
        )
        
        foreach ($value in $valuesToRemove) {
            if (Get-ItemProperty -Path $preferencesPath -Name $value -ErrorAction SilentlyContinue) {
                Remove-ItemProperty -Path $preferencesPath -Name $value -Force
                Write-Host "Removed registry value: $value"
            }
        }
        
        Write-Host "Successfully updated Outlook preferences registry values"
    }
    catch {
        Write-Warning "Failed to update Outlook preferences registry values: $_"
    }
}

# Signature Generation & File Management
function New-SignatureContent {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        [Parameter(Mandatory=$true)]
        [string]$Template,
        [Parameter(Mandatory=$true)]
        [string]$LogoBase64,
        [Parameter(Mandatory=$true)]
        [string]$Address,
        [Parameter(Mandatory=$false)]
        [ValidateSet('New', 'Reply')]
        [string]$Type = "New"
    )

    switch ($Template) {
        'SecondaryDomain' { 
            if ($Type -eq "New") {
                return Get-SecondaryDomainSignatureHTML -user ([PSCustomObject]$UserInfo) -logoBase64 $LogoBase64 -companyAddress $Address
            } else {
                return Get-SecondaryDomainReplySignatureHTML -user ([PSCustomObject]$UserInfo)
            }
        }
        'Alternative' {
            if ($Type -eq "New") {
                return Get-LogoAlternativeSignatureHTML -user ([PSCustomObject]$UserInfo) -logoBase64 $LogoBase64 -companyAddress $Address
            } else {
                return Get-ReplySignatureHTML -user ([PSCustomObject]$UserInfo)
            }
        }
        'Custom' {
            if ($Type -eq "New") {
                return Get-DefaultSignatureHTML -user ([PSCustomObject]$UserInfo) -logoBase64 $LogoBase64 -companyAddress $Address
            } else {
                return Get-ReplySignatureHTML -user ([PSCustomObject]$UserInfo)
            }
        }
        Default {
            if ($Type -eq "New") {
                return Get-DefaultSignatureHTML -user ([PSCustomObject]$UserInfo) -logoBase64 $LogoBase64 -companyAddress $Address
            } else {
                return Get-ReplySignatureHTML -user ([PSCustomObject]$UserInfo)
            }
        }
    }
}

function Get-OutlookSignaturesFolderPath {
    if ($User) { 
        return "C:\Temp\Signatures" 
    } else { 
        $path = "$env:APPDATA\Microsoft\Signatures"
        if (-not [System.IO.Path]::IsPathRooted($path)) {
            $path = [System.IO.Path]::GetFullPath($path)
        }
        return $path
    }
}

# Remove-OldSignatures function with detailed logging
function Remove-OldSignatures {
    param (
        [string]$sigPath,
        [switch]$RemoveAll
    )
    
    try {
        # Create backup folder if it doesn't exist
        $backupPath = Join-Path $env:USERPROFILE "Documents\Outlook Signatures Backup"
        $backupFolder = Join-Path $backupPath (Get-Date -Format "yyyy-MM-dd_HHmmss")
        if (-not (Test-Path $backupPath)) {
            New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
        }
        New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null

        # Backup existing signatures
        Write-Host "Backing up signatures to $backupFolder"
        Get-ChildItem -Path $sigPath -File | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination $backupFolder -Force
        }

        if ($RemoveAll) {
            # Log file cleanup
            $fileCount = (Get-ChildItem -Path $sigPath -File).Count
            LogHelper -logFilePath $logFilePath -message "Found $fileCount signature files to remove" -Level Info -Component "Cleanup" -Action "Files"
            Get-ChildItem -Path $sigPath -File | ForEach-Object {
                LogHelper -logFilePath $logFilePath -message "Removing file: $($_.Name)" -Level Info -Component "Cleanup" -Action "FileRemoval"
                Remove-Item -Path $_.FullName -Force
            }

            # Registry cleanup with detailed logging
            $outlookVersions = @("16.0", "15.0", "14.0")
            $baseProfilePath = "Software\Microsoft\Office"

            foreach ($version in $outlookVersions) {
                $profilesPath = "HKCU:\$baseProfilePath\$version\Outlook\Profiles"
                if (Test-Path $profilesPath) {  
                    LogHelper -logFilePath $logFilePath -message "Processing Outlook $version registry settings" -Level Info -Component "Cleanup" -Action "Registry"
                    
                    # Process signature settings in profiles
                    Get-ChildItem $profilesPath -Recurse | Where-Object { $_.Name -match '9375CFF0413111d3B88A00104B2A6676' } | ForEach-Object {
                        $accountPath = $_.PSPath
                        $account = (Get-ItemProperty -Path $accountPath -ErrorAction SilentlyContinue).'Account Name'
                        LogHelper -logFilePath $logFilePath -message "Cleaning signature settings for account: $account" -Level Info -Component "Cleanup" -Action "RegistryAccount"

                        # Log and clear signature properties
                        $properties = @(
                            @{ Name = "New Signature"; Action = "Reset" },
                            @{ Name = "Reply-Forward Signature"; Action = "Reset" },
                            @{ Name = "ShowAutoSig"; Action = "Remove" },
                            @{ Name = "SignatureAdded"; Action = "Remove" }
                        )

                        foreach ($prop in $properties) {
                            $currentValue = (Get-ItemProperty -Path $accountPath -Name $prop.Name -ErrorAction SilentlyContinue).$($prop.Name)
                            if ($null -ne $currentValue) {
                                LogHelper -logFilePath $logFilePath -message "[$($prop.Action)] $($prop.Name): Current value = '$currentValue'" -Level Info -Component "Cleanup" -Action "RegistryProperty"
                                
                                if ($prop.Action -eq "Reset") {
                                    Set-ItemProperty -Path $accountPath -Name $prop.Name -Value "" -ErrorAction SilentlyContinue
                                } else {
                                    Remove-ItemProperty -Path $accountPath -Name $prop.Name -ErrorAction SilentlyContinue
                                }
                            }
                        }
                    }

                    # Process general Outlook settings
                    $paths = @{
                        "Setup" = @("DisableRoamingSignatures", "DisableRoamingSignaturesTemporaryToggle", "DisableHTMLSignatures", "EnableRoamingSettings")
                        "NewOutlook" = @("EnableHTMLSignatures", "SignatureSyncEnabled")
                    }

                    foreach ($pathKey in $paths.Keys) {
                        $fullPath = "HKCU:\$baseProfilePath\$version\Outlook\$pathKey"
                        if (Test-Path $fullPath) {
                            LogHelper -logFilePath $logFilePath -message "Cleaning $pathKey settings" -Level Info -Component "Cleanup" -Action "RegistrySettings"
                            foreach ($setting in $paths[$pathKey]) {
                                $currentValue = (Get-ItemProperty -Path $fullPath -Name $setting -ErrorAction SilentlyContinue).$setting
                                if ($null -ne $currentValue) {
                                    # Use string format operator
                                    LogHelper -logFilePath $logFilePath -message ("[Remove] {0}: Current value = '{1}'" -f $setting, $currentValue) -Level Info -Component "Cleanup" -Action "RegistryProperty"
                                    Remove-ItemProperty -Path $fullPath -Name $setting -ErrorAction SilentlyContinue
                                }
                            }
                        }
                    }
                }
            }
            LogHelper -logFilePath $logFilePath -message "Registry cleanup completed" -Level Success -Component "Cleanup" -Action "Registry"
        }
    }
    catch {
        LogHelper -logFilePath $logFilePath -message "Failed during cleanup: $_" -Level Error -Component "Cleanup" -Action "Error"
        Write-Warning "Failed during cleanup: $_"
    }
}

# Logging
# LogHelper function with verbose logging
function LogHelper {
    param (
        [string]$logFilePath,
        [string]$message,
        [switch]$Display,
        [ValidateSet('Info', 'Debug', 'Warning', 'Error', 'Success', 'Start', 'End')]
        [string]$Level = 'Info',
        [string]$Component = 'General',
        [string]$Action = ''
    )
    
    # Create detailed log entry
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $username = $env:USERNAME
    $computerName = $env:COMPUTERNAME
    $processId = $PID
    
    $logEntry = [PSCustomObject]@{
        Timestamp = $timestamp
        Level = $Level
        Component = $Component
        Action = $Action
        Message = $message
        User = $username
        Computer = $computerName
        ProcessId = $processId
        ScriptVersion = "1.0"
    }
    
    # Format log message with exact field widths
    $logMessage = "[{0}] [{1,5}] [{2,8}] [{3,10}] - {4} (User: {5} on {6}, PID: {7})" -f `
        $logEntry.Timestamp,
        $logEntry.Level,
        $logEntry.Component,
        $logEntry.Action,
        $logEntry.Message,
        $logEntry.User,
        $logEntry.Computer,
        $logEntry.ProcessId

    # Add separator lines for start/end messages
    if ($Level -in 'Start','End') {
        $separator = "#" * 129
        "$separator`n$logMessage`n$separator" | Out-File -FilePath $logFilePath -Append -Encoding UTF8
    } else {
        $logMessage | Out-File -FilePath $logFilePath -Append -Encoding UTF8
    }
    
    # Console output with color coding
    if ($Display) {
        $color = switch ($Level) {
            'Info'    { 'White' }
            'Debug'   { 'Gray' }
            'Warning' { 'Yellow' }
            'Error'   { 'Red' }
            'Success' { 'Green' }
            'Start'   { 'Cyan' }
            'End'     { 'Cyan' }
            Default   { 'White' }
        }
        Write-Host $logMessage -ForegroundColor $color
    }
}

# Add Template Caching
$script:templateCache = @{}

function Get-CachedTemplate {
    param($templateName, $templateContent)
    
    if (-not $script:templateCache[$templateName]) {
        $script:templateCache[$templateName] = $templateContent
    }
    return $script:templateCache[$templateName]
}

# Add Progress Reporting
function Write-ProgressHelper {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Enhanced Error Management
function Write-ErrorLog {
    param(
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    
    $errorDetails = @{
        Message = $Message
        Exception = $ErrorRecord.Exception.Message
        ScriptStackTrace = $ErrorRecord.ScriptStackTrace
        Category = $ErrorRecord.CategoryInfo.Category
        TimeStamp = (Get-Date).ToString()
    }
    
    $errorJson = $errorDetails | ConvertTo-Json
    Add-Content -Path "$logBasePath\OutlookSignature_$(Get-Date -Format 'yyyyMMdd')_errors.json" -Value $errorJson
}

# Add Resource Cleanup
function Clear-Resources {
    $script:cachedToken = $null
    $script:tokenExpiration = [DateTime]::MinValue
    $script:templateCache.Clear()
    [System.GC]::Collect()
}

# Add Retry Logic for API Calls
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 2
    )
    
    $retryCount = 0
    do {
        try {
            return & $ScriptBlock
        }
        catch {
            $retryCount++
            if ($retryCount -eq $MaxRetries) { throw }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    } while ($retryCount -lt $MaxRetries)
}

# Add Signature Verification
function Test-SignatureFiles {
    param(
        [string]$SignaturePath,
        [string]$EmailAddress
    )
    
    $expectedFiles = @(
        "New ($EmailAddress).htm",
        "Reply ($EmailAddress).htm"
    )
    
    $missingFiles = $expectedFiles | Where-Object {
        -not (Test-Path (Join-Path $SignaturePath $_))
    }
    
    return @{
        Success = $missingFiles.Count -eq 0
        MissingFiles = $missingFiles
    }
}

# Update signature generation function
function New-OutlookSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Logo = "Logo-Default",
        [string]$PrimaryEmail,
        [string]$SecondaryEmail,
        [Parameter(ValueFromPipeline)]
        [hashtable]$UserInfo
    )
    
    begin {
        Write-ProgressHelper -Activity "Creating Outlook Signature" -Status "Initializing" -PercentComplete 0
    }
    
    process {
        try {
            # Determine if the user has a secondary SecondaryDomain mailbox
            $hasSecondaryDomainMailbox = $false
            if ($SecondaryEmail -match $DOMAINS.Secondary) {
                $hasSecondaryDomainMailbox = $true
            }

            # Get template information based on logo and mailbox presence
            Get-SignatureTemplate -Logo $Logo -HasSecondaryDomainMailbox $hasSecondaryDomainMailbox

            # Generate primary email signatures
            $primaryNewFile = Get-SignatureFileName -EmailAddress $PrimaryEmail -Type "New" -Logo $Logo
            $primaryReplyFile = Get-SignatureFileName -EmailAddress $PrimaryEmail -Type "Reply" -Logo $Logo
            
            # Generate primary signatures
            Set-Content -Path "$SignaturePath\$primaryNewFile" -Value (Get-CachedTemplate -templateName $primaryNewFile -templateContent (Get-SignatureTemplate -LogoKey $Logo).New)
            Set-Content -Path "$SignaturePath\$primaryReplyFile" -Value (Get-CachedTemplate -templateName $primaryReplyFile -templateContent (Get-SignatureTemplate -LogoKey $Logo).Reply)
            
            # If secondary SecondaryDomain email exists, generate additional signatures
            if ($hasSecondaryDomainMailbox) {
                $secondaryNewFile = Get-SignatureFileName -EmailAddress $SecondaryEmail -Type "New" -Logo "Logo-SecondaryDomain"
                $secondaryReplyFile = Get-SignatureFileName -EmailAddress $SecondaryEmail -Type "Reply" -Logo "Logo-SecondaryDomain"
                
                # Use SecondaryDomain templates for secondary email
                Set-Content -Path "$SignaturePath\$secondaryNewFile" -Value (Get-CachedTemplate -templateName $secondaryNewFile -templateContent (Get-SignatureTemplate -LogoKey "Logo-SecondaryDomain").New)
                Set-Content -Path "$SignaturePath\$secondaryReplyFile" -Value (Get-CachedTemplate -templateName $secondaryReplyFile -templateContent (Get-SignatureTemplate -LogoKey "Logo-SecondaryDomain").Reply)
            }
        }
        catch {
            Write-ErrorLog -Message "Failed to create signature" -ErrorRecord $_
            throw
        }
    }
    
    end {
        Clear-Resources
        Write-ProgressHelper -Activity "Creating Outlook Signature" -Status "Complete" -PercentComplete 100
    }
}

# Add logging constants
$LOG_SECTIONS = @{
    Header = "=== {0} ==="
    Separator = "#" * 130
}

function Write-LogSection {
    param (
        [string]$Title,
        [hashtable]$Data,
        [string]$LogPath
    )
    
    Add-Content -Path $LogPath -Value ""
    Add-Content -Path $LogPath -Value ($LOG_SECTIONS.Header -f $Title)
    foreach ($key in $Data.Keys) {
        Add-Content -Path $LogPath -Value "$key`: $($Data[$key])"
    }
}

function Write-SignatureLog {
    param (
        [PSCustomObject]$UserInfo,
        [string]$Logo,
        [string]$LogPath
    )
    
    # Start new log entry
    Add-Content -Path $LogPath -Value $LOG_SECTIONS.Separator
    Add-Content -Path $LogPath -Value $LOG_SECTIONS.Separator
    Add-Content -Path $LogPath -Value "[Info] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Script started for user: $($UserInfo.EmailAddress)"
    Add-Content -Path $LogPath -Value "[Info] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Using logo: $Logo"
    
    # Write execution summary
    Write-LogSection -Title "EXECUTION SUMMARY" -Data @{
        "Timestamp" = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        "Script Version" = "1.0"
    } -LogPath $LogPath
    
    # Write user information
    Write-LogSection -Title "USER INFORMATION" -Data @{
        "Display Name" = $UserInfo.DisplayName
        "Job Title" = $UserInfo.JobTitle
        "Primary Email" = $UserInfo.EmailAddress
        "Phone Number" = $UserInfo.TelephoneNumber
        "Office Location" = $UserInfo.OfficeLocation
        "Office Address" = $Addresses[$UserInfo.OfficeLocation]
    } -LogPath $LogPath
    
   # Write signature configuration
   Write-LogSection -Title "SIGNATURE CONFIGURATION" -Data @{
        "Template Used" = if ($Logo -eq "Logo-Alternative") { "Alternative" } elseif ($HasSecondaryDomainMailbox) { "SecondaryDomain" } else { "Default" }
        "Logo Selected" = $Logo
        "Signature Path" = Join-Path $env:APPDATA "Microsoft\Signatures"
    } -LogPath $LogPath

    # Write files created section
    Write-LogSection -Title "FILES CREATED" -Data @{
        "Primary New Email" = Join-Path $env:APPDATA "Microsoft\Signatures\$(Get-SignatureFileName -EmailAddress $UserInfo.EmailAddress -Type 'New' -Logo $Logo)"
        "Primary Reply" = Join-Path $env:APPDATA "Microsoft\Signatures\$(Get-SignatureFileName -EmailAddress $UserInfo.EmailAddress -Type 'Reply' -Logo $Logo)"
    } -LogPath $LogPath

# Write SecondaryDomain details if applicable
if ($HasSecondaryDomainMailbox) {
    Write-LogSection -Title "SECONDARYDOMAIN DETAILS" -Data @{
        "Has SecondaryDomain Account" = $HasSecondaryDomainMailbox
        "Secondary New Email" = Join-Path $env:APPDATA "Microsoft\Signatures\$(Get-SignatureFileName -EmailAddress $SecondaryEmail -Type 'New' -Logo 'Logo-SecondaryDomain')"
        "Secondary Reply" = Join-Path $env:APPDATA "Microsoft\Signatures\$(Get-SignatureFileName -EmailAddress $SecondaryEmail -Type 'Reply' -Logo 'Logo-SecondaryDomain')"
    } -LogPath $LogPath
}
else {
    Write-LogSection -Title "SECONDARYDOMAIN DETAILS" -Data @{
        "Has SecondaryDomain Account" = $false
    } -LogPath $LogPath
}
}
#endregion

#region Signature Templates 
# Template functions
function Get-DefaultSignatureHTML {
    param (
        [PSCustomObject]$user,
        [string]$logoBase64,
        [string]$companyAddress
    )
        
    return @"
<table style="text-align: left; text-indent: 0px; box-sizing: border-box; border-collapse: collapse; border-spacing: 0px; color: inherit; background-color: inherit;" cellspacing="0" cellpadding="0">
<tbody>
<tr style="height: 71.5px;">
<td style="text-align: left; text-indent: 0px; border-right: 1pt solid; padding: 0in 5.4pt; vertical-align: top; width: 150px; height: 71.5px;">
<p style="text-align: center; text-indent: 0px; margin: 0in;"><span style="font-family: Aptos, sans-serif; font-size: 11pt; color: #000000;"> <img id="image_1" src="$logoBase64" alt="" /> </span></p>
</td>
<td style="text-align: left; text-indent: 0px; padding: 0in 5.4pt; vertical-align: top; width: 508.13px; height: 71.5px;">
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Aptos, sans-serif; font-size: 12pt;"><span style="font-family: 'Agency FB', sans-serif; color: #000000;"><strong>$($user.DisplayName)</strong></span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Aptos, sans-serif; font-size: 12pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;"><strong>$($user.JobTitle)</strong></span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Calibri, sans-serif; font-size: 11pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;"> Headquarters: (123) 456-1890 | Mobile: $($user.TelephoneNumber) </span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Calibri, sans-serif; font-size: 11pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: blue;"> <u><a style="color: blue; margin: 0px;" href="http://www.company.com/">www.company.com</a></u> </span> <span style="color: #000000;">&nbsp;|&nbsp;</span> <span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;">$companyAddress</span></p>
</td>
</tr>
</tbody>
</table>
<table style="width: 465px;">
<tbody>
<tr>
<td style="width: 465px; text-align: left; text-indent: 0px; margin: 7.5pt 0in 0in;"><span style="font-family: Calibri, sans-serif; font-size: 6pt; color: #000000;">CONFIDENTIALITY: This communication, including attachments, is for the exclusive use of the addressee(s) and may contain proprietary, confidential, or privileged information.&nbsp; If you are not the intended recipient, any use, copying, disclosure, or distribution or the taking of any action in reliance upon this information is strictly prohibited.&nbsp; If you are not the intended recipient, please notify the sender immediately, delete this communication, and destroy all copies.</span></td>
</tr>
</tbody>
</table>
<p><br /><br /></p>
"@
}

function Get-ReplySignatureHTML {
    param (
        [PSCustomObject]$user
    )
    
    return @"
<div style="font-family: Agency FB, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; color: #000000;"><span style="font-size: 12pt;"> <strong>$($user.DisplayName)</strong> </span> <span style="font-size: 10pt;"> <br /> 
<strong>$($user.JobTitle)</strong> 
<br /> Headquarters: (123) 456-7890 | Mobile: $($user.TelephoneNumber) | </span> <span style="font-size: 10pt; color: #51a7f9;"> <u> <a style="color: #51a7f9;" title="www.company.com" href="http://www.company.com"> www.company.com<br /> </a> </u> </span></div>
"@
}

function Get-SecondaryDomainSignatureHTML {
    param (
        [PSCustomObject]$user,
        [string]$logoBase64,
        [string]$companyAddress
    )
        
    return @"
<p style="text-align: left; text-indent: 0px; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 14pt; color: #51a7f9;"> $($user.DisplayName) </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #000000;"> &nbsp; <img id="_x0000_i1025" style="width: 122px; height: 35px; max-width: 780px; margin-top: 0px; margin-bottom: 0px;" src="$logoBase64" alt="SecondaryDomain Logo" width="122" height="35" /> </span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 8pt; color: #51a7f9;"> $($user.JobTitle.ToUpper()) / 808.833.2225 </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 8pt; color: #666666;"> C </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 8pt; color: #51a7f9;"> / $($user.TelephoneNumber) </span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #51a7f9;"> SECONDARYDOMAIN, INC. </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #00a500;"> $($companyAddress.ToUpper()) </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 8pt; color: #00a500;"> <br /><br /> </span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #51a7f9;"> <em><u> <a style="margin-top: 0px; margin-bottom: 0px;" title="https://www.secondarydomain.com" href="https://www.secondarydomain.com"> www.secondarydomain.com </a> </u></em> </span></p>
<p style="text-align: left; text-indent: 0px; margin: 7.5pt 0in 0in;"><span style="font-family: Calibri, sans-serif; font-size: 6pt; color: #000000;"> CONFIDENTIALITY: This communication, including attachments, is for the exclusive use of the addressee(s) and may contain proprietary, confidential,&nbsp;</span><span style="font-family: Calibri, sans-serif; font-size: 6pt; color: #000000;">or <br />privileged information.&nbsp; If you are not the intended recipient, any use, copying, disclosure, or distribution or the taking of any action in reliance upon this <br />information is strictly prohibited.&nbsp; If you are not the intended recipient, please notify the sender immediately, delete this communication, and destroy all copies. </span></p>
"@
}

function Get-SecondaryDomainReplySignatureHTML {
    param (
        [PSCustomObject]$user
    )
       
    return @"
<p style="text-align: left; text-indent: 0px; background-color: #ffffff; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 14pt; color: #51a7f9;"> $($user.DisplayName) </span></p>
<p style="text-align: left; text-indent: 0px; background-color: #ffffff; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #00a500;"> SECONDARYDOMAIN, INC. </span> <span style="font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #00a500;"> <strong>$($user.JobTitle.ToUpper())</strong> </span></p>
<p style="text-align: left; text-indent: 0px; background-color: #ffffff; margin: 0in;"><span style="font-family: 'Century Gothic', sans-serif; font-size: 8pt; color: #51a7f9;"> C/ $($user.TelephoneNumber) E/ $($user.EmailAddress.ToUpper())<br /> </span></p>
"@
}

function Get-LogoAlternativeSignatureHTML {
    param (
        [PSCustomObject]$user,
        [string]$logoBase64,
        [string]$companyAddress
    )
    
    return @"
<table style="text-align: left; text-indent: 0px; box-sizing: border-box; border-collapse: collapse; border-spacing: 0px; color: inherit; background-color: inherit;" cellspacing="0" cellpadding="0">
<tbody>
<tr style="height: 71.5px;">
<td style="text-align: left; text-indent: 0px; border-right: 1pt solid; padding: 0in 5.4pt; vertical-align: top; width: 150px; height: 71.5px;">
<p style="text-align: center; text-indent: 0px; margin: 0in;"><span style="font-family: Aptos, sans-serif; font-size: 11pt; color: #000000;"> <img id="image_1" src="$logoBase64" alt="" /> </span></p>
</td>
<td style="text-align: left; text-indent: 0px; padding: 0in 5.4pt; vertical-align: top; width: 508.13px; height: 71.5px;">
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Aptos, sans-serif; font-size: 12pt;"><span style="font-family: 'Agency FB', sans-serif; color: #000000;"><strong>$($user.DisplayName)</strong></span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Aptos, sans-serif; font-size: 12pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;"><strong>$($user.JobTitle)</strong></span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Calibri, sans-serif; font-size: 11pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;"> Headquarters: (123) 456-7890 | Mobile: $($user.TelephoneNumber) </span></p>
<p style="text-align: left; text-indent: 0px; margin: 0in; font-family: Calibri, sans-serif; font-size: 11pt;"><span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: blue;"> <u><a style="color: blue; margin: 0px;" href="http://www.company.com/">www.company.com</a></u> </span> <span style="color: #000000;">&nbsp;|&nbsp;</span> <span style="font-family: 'Agency FB', sans-serif; font-size: 10pt; color: #000000;">$companyAddress</span></p>
</td>
</tr>
</tbody>
</table>
<table style="width: 465px;">
<tbody>
<tr>
<td style="width: 465px; text-align: left; text-indent: 0px; margin: 7.5pt 0in 0in;"><span style="font-family: Calibri, sans-serif; font-size: 6pt; color: #000000;">CONFIDENTIALITY: This communication, including attachments, is for the exclusive use of the addressee(s) and may contain proprietary, confidential, or privileged information.&nbsp; If you are not the intended recipient, any use, copying, disclosure, or distribution or the taking of any action in reliance upon this information is strictly prohibited.&nbsp; If you are not the intended recipient, please notify the sender immediately, delete this communication, and destroy all copies.</span></td>
</tr>
</tbody>
</table>
<p><br /><br /></p>
"@
}

# Add Alternative reply signature template if do not want to use the default one
#endregion

#region Main Execution
try {
    # Handle cleanup switch first
    if ($Cleanup) {
        LogHelper -logFilePath $logFilePath -message "Starting signature cleanup process" -Display -Level Start -Component "Cleanup" -Action "Initialize"
        Write-Host "`n=== Cleaning up Signatures ===" -ForegroundColor Cyan
        $signaturesPath = Get-OutlookSignaturesFolderPath
        Remove-OldSignatures -sigPath $signaturesPath -RemoveAll
        LogHelper -logFilePath $logFilePath -message "Signature cleanup completed successfully" -Display -Level End -Component "Cleanup" -Action "Complete"
        Write-Host "Cleanup completed" -ForegroundColor Green
        exit 0
    }

    LogHelper -logFilePath $logFilePath -message "Starting signature creation process" -Display -Level Start -Component "Signature" -Action "Initialize"
    
    # Initialize with detailed logging
    LogHelper -logFilePath $logFilePath -message "Script version: 1.0" -Display -Level Info -Component "Signature" -Action "Version"
    LogHelper -logFilePath $logFilePath -message "Parameters - User: $User, Template: $Template, Logo: $Logo" -Level Info -Component "Signature" -Action "Config"
    
    # Get token and user info with logging
    LogHelper -logFilePath $logFilePath -message "Requesting Graph API access token" -Level Info -Component "Auth" -Action "Token"
    $accessToken = Get-GraphAccessToken -config $authConfig
    LogHelper -logFilePath $logFilePath -message "Token acquired successfully" -Level Success -Component "Auth" -Action "Token"
    
    $userEmail = if ($User) { 
        "$User@PrimaryDomain.com" 
    } else { 
        "$env:username@PrimaryDomain.com" 
    }
    LogHelper -logFilePath $logFilePath -message "Processing for email: $userEmail" -Level Info -Component "Signature" -Action "User"
    
    # User information retrieval with detailed logging
    LogHelper -logFilePath $logFilePath -message "Retrieving user information from Graph API" -Level Info -Component "Graph" -Action "UserInfo"
    $userInfo = Get-GraphUserInfo -userPrincipalName $userEmail -accessToken $accessToken
    LogHelper -logFilePath $logFilePath -message "Retrieved info for: $($userInfo.displayName)" -Level Success -Component "Graph" -Action "UserInfo"
    
    # Directory setup with logging
    $sigPath = Get-OutlookSignaturesFolderPath
    LogHelper -logFilePath $logFilePath -message "Using signature path: $sigPath" -Level Info -Component "Signature" -Action "Path"
    if (-not (Test-Path $sigPath)) {
        LogHelper -logFilePath $logFilePath -message "Creating signature directory" -Level Info -Component "Signature" -Action "Directory"
        New-Item -ItemType Directory -Path $sigPath -Force | Out-Null
    }
    
    # Template selection logging
    $templateType = if ($Template) {
        $Template
    } else {
        switch ($Logo) {
            "Logo-Alternative" { "Alternative" }
            "Logo-SecondaryDomain" { "SecondaryDomain" }
            Default { "Default" }
        }
    }
    LogHelper -logFilePath $logFilePath -message "Selected template type: $templateType" -Level Info -Component "Signature" -Action "Template"
    
    # Logo selection logging
    $logoBase64 = if ($Base64Logos[$Logo]) {
        $Base64Logos[$Logo]
    } else {
        $Base64Logos["Logo-Default"]
    }
    LogHelper -logFilePath $logFilePath -message "Using logo type: $Logo" -Level Info -Component "Signature" -Action "Logo"
    
    # Get office address based on user's location
    $address = $Addresses[$userInfo.OfficeLocation]
    LogHelper -logFilePath $logFilePath -message "Using office address: $address" -Level Info -Component "Signature" -Action "Address"

    # Generate signatures with detailed logging
    LogHelper -logFilePath $logFilePath -message "Generating signature content" -Level Info -Component "Signature" -Action "Generate"
    $signatures = @(
        @{ 
            Type = "New"
            Content = New-SignatureContent `
                -UserInfo $userInfo `
                -Template $templateType `
                -LogoBase64 $logoBase64 `
                -Address $address `
                -Type "New"
        }
        @{ 
            Type = "Reply"
            Content = New-SignatureContent `
                -UserInfo $userInfo `
                -Template $templateType `
                -LogoBase64 $logoBase64 `
                -Address $address `
                -Type "Reply"
        }
    )
    
    # Save signatures with logging
    foreach ($sig in $signatures) {
        try {
            $fileName = Get-SignatureFileName -EmailAddress $userInfo.EmailAddress -Type $sig.Type -Logo $Logo
            LogHelper -logFilePath $logFilePath -message "Creating signature file: $fileName" -Level Info -Component "Signature" -Action "Create"
            $sig.Content | Out-File -FilePath (Join-Path $sigPath $fileName) -Encoding UTF8
            LogHelper -logFilePath $logFilePath -message "Successfully created: $fileName" -Level Success -Component "Signature" -Action "Create"
        }
        catch {
            LogHelper -logFilePath $logFilePath -message "Failed to create $fileName : $_" -Level Error -Component "Signature" -Action "Create"
            throw
        }
    }

    # Registry updates with detailed logging
    if (-not $User) {
        LogHelper -logFilePath $logFilePath -message "Updating Outlook registry settings" -Level Info -Component "Registry" -Action "Begin"
        try {
            Update-OutlookRegistry `
                -newEmailSignature (Get-SignatureFileName -EmailAddress $userInfo.EmailAddress -Type "New" -Logo $Logo) `
                -replySignature (Get-SignatureFileName -EmailAddress $userInfo.EmailAddress -Type "Reply" -Logo $Logo) `
                -userEmail $userInfo.EmailAddress
            
            # Set Outlook preferences
            Set-OutlookPreferences
            
            LogHelper -logFilePath $logFilePath -message "Registry settings updated successfully" -Level Success -Component "Registry" -Action "Complete"
        }
        catch {
            LogHelper -logFilePath $logFilePath -message "Failed to update registry: $_" -Level Error -Component "Registry" -Action "Error"
            throw
        }
    }

    # SecondaryDomain account handling with detailed logging
    LogHelper -logFilePath $logFilePath -message "Checking for SecondaryDomain account" -Level Info -Component "SecondaryDomain" -Action "Check"
    $secondarydomainAccount = Get-SecondaryDomainUserInfo -userPrincipalName $userEmail -accessToken $accessToken
    
    if ($secondarydomainAccount.HasAccount) {
        LogHelper -logFilePath $logFilePath -message "Creating SecondaryDomain signature for $($secondarydomainAccount.Email)" -Level Info -Component "SecondaryDomain" -Action "Create"
        
        # Validate Logo-SecondaryDomain logo exists and fallback to default if not
        $secondarydomainLogo = if ($Base64Logos["Logo-SecondaryDomain"] -and -not [string]::IsNullOrWhiteSpace($Base64Logos["Logo-SecondaryDomain"])) {
            $Base64Logos["Logo-SecondaryDomain"]
        } else {
            LogHelper -logFilePath $logFilePath -message "Logo-SecondaryDomain logo not found, using default logo" -Level Warning -Component "SecondaryDomain" -Action "Logo"
            $Base64Logos["Logo-Default"]
        }
        
        $secondarydomainNewFile = "New_DOMAIN2 ($($secondarydomainAccount.Email)).htm"
        $secondarydomainReplyFile = "Reply_DOMAIN2 ($($secondarydomainAccount.Email)).htm"
        
        $secondarydomainNewPath = Join-Path $sigPath $secondarydomainNewFile
        $secondarydomainReplyPath = Join-Path $sigPath $secondarydomainReplyFile

        # Generate SecondaryDomain signatures with validated logo
        $secondarydomainEmailContent = New-SignatureContent `
            -UserInfo $userInfo `
            -Template "SecondaryDomain" `
            -LogoBase64 $secondarydomainLogo `
            -Address $address `
            -Type "New"
            
        $secondarydomainReplyContent = New-SignatureContent `
            -UserInfo $userInfo `
            -Template "SecondaryDomain" `
            -LogoBase64 $secondarydomainLogo `
            -Address $address `
            -Type "Reply"
        
        $secondarydomainEmailContent | Out-File -FilePath $secondarydomainNewPath -Encoding UTF8
        $secondarydomainReplyContent | Out-File -FilePath $secondarydomainReplyPath -Encoding UTF8
        
        LogHelper -logFilePath $logFilePath -message "SecondaryDomain signature creation completed" -Level Success -Component "SecondaryDomain" -Action "Complete"
    }
    else {
        LogHelper -logFilePath $logFilePath -message "No SecondaryDomain account found" -Level Info -Component "SecondaryDomain" -Action "Skip"
    }
    
    # Completion logging
    LogHelper -logFilePath $logFilePath -message "Signature creation process completed successfully" -Display -Level End -Component "Signature" -Action "Complete"

    # After all signature operations complete, display formatted information
    Write-Host "`n[Info] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Script started for user: $($userInfo.EmailAddress)"
    Write-Host "[Info] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Using logo: $Logo`n"
    
    Write-Host "=== EXECUTION SUMMARY ==="
    Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host "Script Version: 1.0.0`n"
    
    Write-Host "=== USER INFORMATION ==="
    Write-Host "Office Location: $($userInfo.OfficeLocation)"
    Write-Host "Job Title: $($userInfo.JobTitle)"
    Write-Host "Phone Number: $($userInfo.TelephoneNumber)"
    Write-Host "Primary Email: $($userInfo.EmailAddress)"
    Write-Host "Office Address: $($Addresses[$userInfo.OfficeLocation])"
    Write-Host "Display Name: $($userInfo.DisplayName)`n"
    
    Write-Host "=== SIGNATURE CONFIGURATION ==="
    Write-Host "Logo Selected: $Logo"
    Write-Host "Signature Path: $sigPath"
    Write-Host "Template Used: $templateType`n"
    
    Write-Host "=== FILES CREATED ==="
    Write-Host "Primary Reply: $(Join-Path $sigPath (Get-SignatureFileName -EmailAddress $userInfo.EmailAddress -Type 'Reply' -Logo $Logo))"
    Write-Host "Primary New Email: $(Join-Path $sigPath (Get-SignatureFileName -EmailAddress $userInfo.EmailAddress -Type 'New' -Logo $Logo))"
    if ($secondarydomainAccount.HasAccount) {
        Write-Host "SecondaryDomain Reply: $(Join-Path $sigPath (Get-SignatureFileName -EmailAddress $secondarydomainAccount.Email -Type 'Reply' -Logo 'Logo-SecondaryDomain'))"
        Write-Host "SecondaryDomain New Email: $(Join-Path $sigPath (Get-SignatureFileName -EmailAddress $secondarydomainAccount.Email -Type 'New' -Logo 'Logo-SecondaryDomain'))"
    }
    Write-Host ""
    
    Write-Host "=== SECONDARYDOMAIN DETAILS ==="
    Write-Host "Has SecondaryDomain Account: $($secondarydomainAccount.HasAccount)"
} catch {
    LogHelper -logFilePath $logFilePath -message "Critical error during signature creation: $($_.Exception.Message)" -Display -Level Error -Component "Signature" -Action "Error"
    LogHelper -logFilePath $logFilePath -message $_.ScriptStackTrace -Level Error -Component "Signature" -Action "StackTrace"
    throw
}
finally {
    LogHelper -logFilePath $logFilePath -message "Cleaning up resources" -Level Info -Component "Cleanup" -Action "Resources"
    Clear-Resources
}
#endregion