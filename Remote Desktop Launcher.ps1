<#############################################
Written by: Paul Willis
Revision date: 8/23/2022

Description: 
- Automates RDP connections to targeted SCCM device collection via local account password.
- Can connect with single computer hostname or multiple
- Runs up to five open sessions, iterates through device list 1 at a time when a session is closed.

#############################################>

#Set the name of default local admin
$LocalUserAccount = "REPLACE_WITH_LAPS_USERNAME"
#Set max number of open remote desktop sessions at a time
$MaxSyncSessions = 5

Import-Module ConfigurationManager

Set-Location CMA:

Add-Type -assembly System.Windows.Forms

#Define Windows Form controls first, logic below
$main_form = New-Object 'System.Windows.Forms.Form' -Property @{TopMost = $true }
$main_form.Text = "Remote Desktop Launcher"
$main_form.ShowIcon=$False
$main_form.AutoSize = $false
$main_form.Height = 640
$main_form.Width = 500
$main_form.FormBorderStyle = "FixedDialog"
$main_form.StartPosition = "CenterScreen"
$main_form.Padding = 30
$main_form.MaximizeBox = $false

$Spacer3 = New-Object System.Windows.Forms.Label 
$Spacer3.Text = ""
$Spacer3.Width = 450
$Spacer3.Dock = "Top"
$Spacer3.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
$Spacer3.Top = 10

$DeviceTargetLabel = New-Object System.Windows.Forms.Label 
$DeviceTargetLabel.Text = "Enter Device Target"
$DeviceTargetLabel.Width = 450
$DeviceTargetLabel.Dock = "Top"
$DeviceTargetLabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
$DeviceTargetLabel.Top = 10

$DeviceTargetTextBox = New-Object System.Windows.Forms.TextBox 
$DeviceTargetTextBox.Text = ""
$DeviceTargetTextBox.TabIndex = 0
$DeviceTargetTextBox.Width = 450
$DeviceTargetTextBox.Dock = "Top"
$DeviceTargetTextBox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$DeviceTargetTextBox.Top = 10
$DeviceTargetTextBox.Select()

$Spacer1 = New-Object System.Windows.Forms.Label 
$Spacer1.Text = ""
$Spacer1.Width = 450
$Spacer1.Dock = "Top"
$Spacer1.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
$Spacer1.Top = 10

$SelectTargetType = New-Object system.Windows.Forms.ComboBox
$SelectTargetType.text = “”
$SelectTargetType.TabIndex = 1
$SelectTargetType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$SelectTargetType.width = 200
$SelectTargetType.Dock = "Top"
@(‘Single Hostname’,’CM Collection’,’CSV of Hostnames (no header)’) | ForEach-Object {[void] $SelectTargetType.Items.Add($_)}
$SelectTargetType.SelectedIndex = 0
$SelectTargetType.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$SelectTargetType.Top = 10

$SelectAdminType = New-Object system.Windows.Forms.ComboBox
$SelectAdminType.text = “”
$SelectAdminType.TabIndex = 2
$SelectAdminType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$SelectAdminType.width = 200
$SelectAdminType.Dock = "Top"
@(‘LAPS Account’,’SYS Account’) | ForEach-Object {[void] $SelectAdminType.Items.Add($_)}
$SelectAdminType.SelectedIndex = 0
$SelectAdminType.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$SelectAdminType.Top = 10

$Spacer2 = New-Object System.Windows.Forms.Label 
$Spacer2.Text = ""
$Spacer2.Width = 450
$Spacer2.Dock = "Top"
$Spacer2.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
$Spacer2.Top = 10

$LaunchButton = New-Object System.Windows.Forms.Button
$LaunchButton.Top = 10
$LaunchButton.TabIndex = 3
$LaunchButton.Dock = "Top"
$LaunchButton.Size = New-Object System.Drawing.Size(100,29)
$LaunchButton.Text = "Launch Remote Desktop Connections"
$LaunchButton.Top = 10
$main_form.AcceptButton = $LaunchButton

$LogLabel = New-Object System.Windows.Forms.Label 
$LogLabel.Text = "Connections Log"
$LogLabel.Width = 450
$LogLabel.Dock = "Top"
$LogLabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
$LogLabel.Top = 10

$LogTextBox = New-Object System.Windows.Forms.TextBox
$LogTextBox.Location  = New-Object System.Drawing.Point(10,10)
$LogTextBox.Width = 450
$LogTextBox.Multiline = "True"
$LogTextBox.Scrollbars = "Vertical"
$LogTextBox.Height = 320
$LogTextBox.Dock = "Top"
$LogTextBox.ReadOnly = "True"
$LogTextBox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$LogTextBox.Top = 10
$LogTextBox.BackColor = "White"

$main_form.Controls.Add($LaunchButton)
$main_form.Controls.Add($Spacer2)
$main_form.Controls.Add($SelectAdminType)
$main_form.Controls.Add($SelectTargetType)
$main_form.Controls.Add($Spacer1)
$main_form.Controls.Add($DeviceTargetTextBox)
$main_form.Controls.Add($DeviceTargetLabel)
$main_form.Controls.Add($Spacer3)
$main_form.Controls.Add($LogTextBox)
$main_form.Controls.Add($LogLabel)


$LaunchButton.Add_Click({

    $Computers

    #Lock input from changes during script execution
    $DeviceTargetTextBox.ReadOnly = "True"

    #Clear cache credentials
    cmdkey /list | ForEach-Object{if($_ -like "*Target:*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}}

    #Depending on selection:
    #Get computers from MECM device collection
    if($SelectTargetType.Text -eq "CM Collection"){
        $LogTextBox.Text += "`r`n`r`nFetching collection..."
        $LogTextBox.SelectionStart = $LogTextBox.Text.Length
        $LogTextBox.ScrollToCaret()
        $coll = Get-CMDeviceCollection -Name $DeviceTargetTextBox.Text
        Start-Sleep -Seconds 10
        $LogTextBox.Text += "`r`nFetching devices from collection..."
        $LogTextBox.SelectionStart = $LogTextBox.Text.Length
        $LogTextBox.ScrollToCaret()
        $Computers = Get-CMDevice -Collection $coll -Resource | Select-Object Name, Active, IPAddresses
    }
    #Get single computer from the input
    elseif ($SelectTargetType.Text -eq "Single Hostname") {
        $Computers = @{}
        $Computers.Add("Name", $DeviceTargetTextBox.Text)
    } 
    #Get computers from CSV as file path
    elseif ($SelectTargetType.Text -eq "CSV of Hostnames (no header)") {
        $Computers = Import-CSV $DeviceTargetTextBox.Text.Replace('"',"") -Header "Name"
    }
    
    #Retrieve active directory info for LAPS password
    $LogTextBox.Text += "`r`nGetting device info...`r`n"
    $LogTextBox.SelectionStart = $LogTextBox.Text.Length
    $LogTextBox.ScrollToCaret()
    $Computers | ForEach-Object {
            $psw = Get-ADComputer -Identity $_.Name -Properties ms-Mcs-AdmPwd, ms-Mcs-AdmPwdExpirationTime
            $psw | Sort-Object ms-Mcs-AdmPwdExpirationTime | Format-Table -AutoSize Name, DnsHostName, ms-Mcs-AdmPwd, ms-Mcs-AdmPwdExpirationTime
            $_ | Add-Member -NotePropertyName LAPS -NotePropertyValue $psw."ms-Mcs-AdmPwd"
            if ($_.IPAddresses) {
                $_.IPAddresses = $_.IPAddresses[0]
            }

            #Note: Currently unused data can be added to the connection log (i.e. LAPS expiration, IP address, etc.)
            $LogTextBox.Text += "`r`n" + $_.Name + " " + $_.LAPS
            $LogTextBox.SelectionStart = $LogTextBox.Text.Length
            $LogTextBox.ScrollToCaret()
    }

    #Temporarily suspend the RDP untrusted client prompt, remove if this is not an issue for your system
    Set-ItemProperty -Path “HKCU:\Software\Microsoft\Terminal Server Client” -name “AuthenticationLevelOverride” -value 0

    #Main connection script
    function Connect-RDP {

      param (
        [Parameter(Mandatory=$true)]
        $ComputerList
      )

      #If SYS connection option is selected, prompt for credentials
      $Credential
      if ($SelectAdminType.Text -eq "SYS Account"){
        $Credential = $Host.ui.PromptForCredential("Need credentials", "Please enter your user name and password.", "", "NetBiosUserName")
      }
      
      $LogTextBox.Text += "`r`n`r`nLaunching Remote Desktop..."
      $LogTextBox.SelectionStart = $LogTextBox.Text.Length
      $LogTextBox.ScrollToCaret()

      #Iterate through computers
      $ComputerList | ForEach-Object {
        #When maximum sessions are open, wait for a remote desktop to be closed before opening the next
        while ((get-process -ea silentlycontinue mstsc).count -ge $MaxSyncSessions) {
                Start-Sleep -Seconds 3
        }

        $Hostname = ""
        if ($_.IPAddresses){
            $Hostname = $_.IPAddresses
        } else {
            $Hostname = $_.Name
        }

        #Delete any existing cache cred for this RDP connection
        cmdkey /delete:$Hostname

        #Depending on selection, create credential with local account and LAPS password
        if ($SelectAdminType.Text -eq "LAPS Account"){
            $Password = ConvertTo-SecureString $_.LAPS -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential ("$Hostname\$LocalUserAccount", $password)
            $LogTextBox.Text += $Hostname+"\"+$LocalAccountName + "  " + $Password
            $LogTextBox.Text += "`r`n" + $Credential.UserName + " " + $Credential.GetNetworkCredential().Password
        }

        #Retrieve User and Password from credential object and create cache credential
        $User = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password
        cmdkey.exe /generic:$Hostname /user:$User /pass:$Password
        $LogTextBox.Text += "`r`nConnecting to " + $_.Name

        #Settings for remote desktop launch... more options here: 
        #https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/mstsc
        $arguments = "/v $Hostname /w:1422 /h:800"
        #Note: Could be extended to use a customized RDP config for granular control:
        #https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/rdp-files

        #If there are already any session open launch RDP in the background
        if ((get-process -ea silentlycontinue mstsc).count -ge 1) {
            Start-Process mstsc.exe -ArgumentList $arguments -WindowStyle Minimized
        } 
        #Else, if this is the first session, launch normally
        else {
            Start-Process mstsc.exe -ArgumentList $arguments
        }
        Start-Sleep -Milliseconds 500
      }
      $LogTextBox.Text += "`r`n`r`nEnd of the Device List"
      $LogTextBox.SelectionStart = $LogTextBox.Text.Length
      $LogTextBox.ScrollToCaret()

      #Enable input after computer loop has finished
      $DeviceTargetTextBox.ReadOnly = "False"
    }

    #Pass computer list to main connection script above
    Connect-RDP -ComputerList $Computers

    #Removing the temporary "untrusted client prompt" suspension, remove this (and Set-ItemProperty above) if it is not an issue for your system
    Remove-ItemProperty -Path “HKCU:\Software\Microsoft\Terminal Server Client” -name “AuthenticationLevelOverride”

    Start-Sleep -seconds 4

    #Note: Third instance of cleaning cache credentials... debug to simplify?
    $Computers | ForEach-Object {
        $Hostname = $_.IPAddresses
        cmdkey /delete:$Hostname
    }
})

$main_form.ShowDialog()