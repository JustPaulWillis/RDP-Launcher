[cmdletbinding(SupportsShouldProcess=$True)]

Param(
)

Function Get-VIAOSVersion([ref]$OSv)
{
    $OS = Get-WmiObject -Class Win32_OperatingSystem
    
    Switch -Regex ($OS.Version)
    {
    "6.1"
        {If ($OS.ProductType -eq 1)
            {$OSv.value = "Windows 7 SP1"}
                Else
            {$OSv.value = "Windows Server 2008 R2"}
        }
    "6.2"
        {If ($OS.ProductType -eq 1)
            {$OSv.value = "Windows 8"}
                Else
            {$OSv.value = "Windows Server 2012"}
        }
    "6.3"
        {If ($OS.ProductType -eq 1)
            {$OSv.value = "Windows 8.1"}
                Else
            {$OSv.value = "Windows Server 2012 R2"}
        }
    "10"
        {If ($OS.ProductType -eq 1)
            {$OSv.value = "Windows 10"}
                Else
            {$OSv.value = "Windows Server 2016"}
        }
    DEFAULT { "Version not listed" }
    } 
}

Function Import-VIASMSTSENV
{
    try
    {
        $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
        Write-Output "$ScriptName - tsenv is $tsenv "
        $MDTIntegration = $true
        
        #$tsenv.GetVariables() | % { Write-Output "$ScriptName - $_ = $($tsenv.Value($_))" }
    }
    catch
    {
        Write-Output "$ScriptName - Unable to load Microsoft.SMS.TSEnvironment"
        Write-Output "$ScriptName - Running in standalonemode"
        $MDTIntegration = $false
    }
    Finally
    {
        if ($MDTIntegration -eq $true)
        {
            $Logpath = $tsenv.Value("LogPath")
            $LogFile = $Logpath + "\" + "$ScriptName.txt"
        }
    Else{
            $Logpath = $env:TEMP
            $LogFile = $Logpath + "\" + "$ScriptName.txt"
        }
    }
    Return $MDTIntegration
}
Function Start-VIALogging
{
    Start-Transcript -path $LogFile -Force
}
Function Stop-VIALogging
{
    Stop-Transcript
}
Function Invoke-VIAExe
{
    [CmdletBinding(SupportsShouldProcess=$true)]

    param(
        [parameter(mandatory=$true,position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Executable,

        [parameter(mandatory=$false,position=1)]
        [string]
        $Arguments
    )

    if ($Arguments -eq "")
    {
        Write-Verbose "Running Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru"
        $ReturnFromEXE = Start-Process -FilePath $Executable -NoNewWindow -Wait -Passthru
    }
    else
    {
        Write-Verbose "Running Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru"
        $ReturnFromEXE = Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru
    }
   
    Write-Verbose "Returncode is $($ReturnFromEXE.ExitCode)"
    Return $ReturnFromEXE.ExitCode
}
Function Invoke-VIAMsi
{
    [CmdletBinding(SupportsShouldProcess=$true)]

    param(
        [parameter(mandatory=$true,position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $MSI,

        [parameter(mandatory=$false,position=1)]
        [string]
        $Arguments
    )

    #Set MSIArgs
    $MSIArgs = "/i " + $MSI + " " + $Arguments

    if ($Arguments -eq "")
    {
        $MSIArgs = "/i " + $MSI
    }
    else
    {
        $MSIArgs = "/i " + $MSI + " " + $Arguments
    }
    
    Write-Verbose "Running Start-Process -FilePath msiexec.exe -ArgumentList $MSIArgs -NoNewWindow -Wait -Passthru"
    $ReturnFromEXE = Start-Process -FilePath msiexec.exe -ArgumentList $MSIArgs -NoNewWindow -Wait -Passthru
    
    Write-Verbose "Returncode is $($ReturnFromEXE.ExitCode)"
    Return $ReturnFromEXE.ExitCode
}
Function Invoke-VIAMsu
{
    [CmdletBinding(SupportsShouldProcess=$true)]

    param(
        [parameter(mandatory=$true,position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $MSU,

        [parameter(mandatory=$false,position=1)]
        [string]
        $Arguments
    )

        #Set MSIArgs
    $MSUArgs = $MSU + " " + $Arguments

    if ($Arguments -eq "")
    {
        $MSUArgs = $MSU  
    }
    else
    {
        $MSUArgs = $MSU + " " + $Arguments
    }

    Write-Verbose "Running Start-Process -FilePath wusa.exe -ArgumentList $MSUArgs -NoNewWindow -Wait -Passthru"
    $ReturnFromEXE = Start-Process -FilePath wusa.exe -ArgumentList $MSUArgs -NoNewWindow -Wait -Passthru
    
    Write-Verbose "Returncode is $($ReturnFromEXE.ExitCode)"
    Return $ReturnFromEXE.ExitCode
}

# Set Vars
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptName = Split-Path -Leaf $MyInvocation.MyCommand.Path
#[xml]$Settings = Get-Content "$ScriptDir\Settings.xml"
$SOURCEROOT = "$SCRIPTDIR\Source"
$LANG = (Get-Culture).Name
$OSV = $Null
$ARCHITECTURE = $env:PROCESSOR_ARCHITECTURE

#Try to Import SMSTSEnv
. Import-VIASMSTSENV

#Start Transcript Logging
. Start-VIALogging

#Detect current OS Version
. Get-VIAOSVersion -osv ([ref]$osv) 

#Output more info
Write-Output ""
Write-Output "$ScriptName - ScriptDir: $ScriptDir"
Write-Output "$ScriptName - SourceRoot: $SOURCEROOT"
Write-Output "$ScriptName - ScriptName: $ScriptName"
Write-Output "$ScriptName - OS Name: $osv"
Write-Output "$ScriptName - OS Architecture: $ARCHITECTURE"
Write-Output "$ScriptName - Current Culture: $LANG"
Write-Output "$ScriptName - Integration with MDT(LTI/ZTI): $MDTIntegration"
Write-Output "$ScriptName - Log: $LogFile"

#Generate more info
if ($MDTIntegration -eq "YES")
{
    $TSMake = $tsenv.Value("Make")
    $TSModel = $tsenv.Value("Model")
    $TSMakeAlias = $tsenv.Value("MakeAlias")
    $TSModelAlias = $tsenv.Value("ModelAlias")
    $TSOSDComputerName = $tsenv.Value("OSDComputerName")
    
    Write-Output "$ScriptName - Make:: $TSMake"
    Write-Output "$ScriptName - Model: $TSModel"
    Write-Output "$ScriptName - MakeAlias: $TSMakeAlias"
    Write-Output "$ScriptName - ModelAlias: $TSModelAlias"
    Write-Output "$ScriptName - OSDComputername: $TSOSDComputerName"
}

#Custom Code Starts--------------------------------------

#Create directory and copy source file into it
New-Item -Path 'C:\Program Files\Remote Desktop Launcher' -ItemType Directory
Write-Output "$ScriptName - Copying source files"
Copy-Item -Path "$ScriptDir\Remote Desktop Launcher.ps1" -Destination 'C:\Program Files\Remote Desktop Launcher'

#Create shortcut on public desktop
Write-Output "$ScriptName - Create a shortcut..."
$TargetFile = "powershell.exe"
$ShortcutFile = "$env:Public\Desktop\Launch Remote Desktop.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.IconLocation="C:\windows\system32\Shell32.dll,17"
$Shortcut.Arguments = "-WindowStyle Hidden -noexit -ExecutionPolicy Bypass -File ""C:\Program Files\Remote Desktop Launcher\Remote Desktop Launcher.ps1"""
$Shortcut.WorkingDirectory = "C:\WINDOWS\system32"
$Shortcut.Save()

#Force shortcut to run as Administrator (will prompt for credentials for non-elevated users)
Write-Output "$ScriptName - Convert shortcut to require Administrator elevation"
$bytes = [System.IO.File]::ReadAllBytes("$env:Public\Desktop\Launch Remote Desktop.lnk")
$bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
[System.IO.File]::WriteAllBytes("$env:Public\Desktop\Launch Remote Desktop.lnk", $bytes)

Stop-VIALogging