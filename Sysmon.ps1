$TranscriptPath = "C:\Windows\Sysmon-Transcript.txt"

if (Test-Path -Path $TranscriptPath) {
    Remove-Item -Path $TranscriptPath -Force
}

Start-Transcript -Path $TranscriptPath -Append


# Which version of Sysmon to install?
# If Sysmon-Test.txt exists, install the Testing version.
if (Test-Path -Path "C:\Windows\Sysmon-Test.txt") {
    $Stage = 'Testing'
}
else {
    $Stage = 'Deployed'
}


# Server with files, used to check for connectivity before downloading
$SourceServer = 'myfileserver'

# Set paths
$ScriptSource = "\\$($SourceServer)\sysmon\Sysmon.ps1"
$Script = "C:\Windows\Sysmon.ps1"

$SysmonConfigSource = "\\$($SourceServer)\sysmon\configs\$($Stage).xml"
$SysmonConfig = 'C:\Windows\Sysmon.xml'

# Set executable paths
if ((Test-Path -Path "C:\Program Files (x86)")) {
    # System is 64 bit
    $SysmonExe = 'C:\Windows\sysmon64.exe'
    $SysmonSource = "\\$($SourceServer)\sysmon\Versions\$($Stage)\sysmon64.exe"
    $SysmonService = 'Sysmon64'
}
else {
    # 32 bit
    $SysmonExe = 'C:\Windows\sysmon.exe'
    $SysmonSource = "\\$($SourceServer)\sysmon\Versions\$($Stage)\sysmon.exe"
    $SysmonService = 'Sysmon'
}

Write-Output "`nVariables"
Write-Output "Stage: $($Stage)"
Write-Output "SysmonConfigSource: $($SysmonConfigSource)"
Write-Output "SysmonConfig: $($SysmonConfig)"
Write-Output "SysmonSource: $($SysmonSource)"
Write-Output "SysmonExe: $($SysmonExe)"
Write-Output "SysmonService: $($SysmonService)"

# Determines if updates are needed to files
$Update = @{
    'Config' = $false
    'Sysmon' = $false
    'Script' = $false
}

# Check if Sysmon is currently installed
$SysmonInstalled = $false
if ((Get-Service -Name $SysmonService -ErrorAction SilentlyContinue)) {
    $SysmonInstalled = $true
    Write-Output "`nSysmon is currently installed."
}

# Check if updates are needed to exe/config
if (Test-Connection $SourceServer -Quiet) {
    Write-Output "`nSource server is available. Checking for updated files."
    if ((Test-Path -Path $SysmonExe)) {
        $ExeCurrentHash = Get-FileHash -Path $SysmonExe
        $ExeSourceHash = Get-FileHash -Path $SysmonSource

        Write-Output "Sysmon Source Hash: $($ExeSourceHash.Hash)"
        Write-Output "Sysmon Current Hash: $($ExeCurrentHash.Hash)"
    
        if ($ExeCurrentHash.Hash -ne $ExeSourceHash.Hash) {
            Write-Output "New Sysmon service available."
            $Update.Sysmon = $true
        }
    }
    else {
        Write-Output "Existing Sysmon service executable not found."
        $Update.Sysmon = $true
    }
    
    if ((Test-Path -Path $SysmonConfig)) {
        $ConfigCurrentHash = Get-FileHash -Path $SysmonConfig
        $ConfigSourceHash = Get-FileHash -Path $SysmonConfigSource

        Write-Output "Config Source Hash: $($ConfigSourceHash.Hash)"
        Write-Output "Config Current Hash: $($ConfigCurrentHash.Hash)"
    
        if ($ConfigCurrentHash.Hash -ne $ConfigSourceHash.Hash) {
            Write-Output "New Sysmon configuration available."
            $Update.Config = $true
        }
    }
    else {
        Write-Output "Existing Sysmon configuration file not found."
        $Update.Config = $true
    }

    if ((Test-Path -Path $Script)) {
        $ScriptCurrentHash = Get-Filehash -Path $Script
        $ScriptSourceHash = Get-FileHash -Path $ScriptSource

        Write-Output "Script Source Hash: $($ScriptSourceHash.Hash)"
        Write-Output "Script Current Hash: $($ScriptCurrentHash.Hash)"

        if ($ScriptCurrentHash.Hash -ne $ScriptSourceHash.Hash) {
            Write-Output "New Sysmon script available."
            $Update.Script = $true
        }
    }
    else {
        Write-Output "Existing script file not found."
        $Update.Script = $true
    }
}
else {
    Write-Output "`nUnable to check for Sysmon updates. $($SourceServer) not reachable."
}

Write-Output "`nUpdates needed?"
Write-Output "Sysmon: $($Update.Sysmon)"
Write-Output "Config: $($Update.Config)"
Write-Output "Script: $($Update.Script)"

if ($Update.Script) {
    Write-Output "Updating script."
    # You cannot use Copy-Item to overwrite the currently running file. 
    # Set-Content *can* overwrite the currently running file. 
    Get-Content -Path $ScriptSource -Raw | Set-Content -NoNewline -Path $Script
}

if ($Update.Sysmon) {
    Write-Output "Updating Sysmon executable."
    if ($SysmonInstalled) {
        # Uninstall sysmon prior to updating, otherwise we will be unable to replace the file.
        Write-Output "Uninstalling existing Sysmon service."
        & $SysmonExe -u
        $SysmonInstalled = $false
    }
    Copy-Item -Path $SysmonSource -Destination $SysmonExe -Force -ErrorAction SilentlyContinue
}

if ($Update.Config) {
    Write-Output "Updating configuration."
    Copy-Item -Path $SysmonConfigSource -Destination $SysmonConfig -Force -ErrorAction SilentlyContinue
}

# Check if service is installed
if (-Not $SysmonInstalled) {
    # Service is not installed

    if (Test-Path -Path $SysmonExe) {
        Write-Output "Sysmon not currently installed. Installing Sysmon service."
        & $SysmonExe -accepteula -i $SysmonConfig
    }
    else {
        Write-Output "Unable to install Sysmon service, $($SysmonExe) not found."
    }
}
else {
    Write-Output "Sysmon service is already installed."
    if ($Update.Config) {
        Write-Output "Updating Sysmon running configuration."
        & $SysmonExe -c $SysmonConfig
    }
}

Stop-Transcript
