<#
Run Command
Does: Deploy CCS Agent on Windows servers from import Csv file or deploy on Full Subscription.
		Csv Import file must contains these headers: 'Name';'ResourceGroupName';'SubscriptionName'
13/09/2019 olivier.poissonnet@dxc.com
changelog:
- v7: add check on CloudOps tag when creating vms variable.
- v6: check prerequisites: PowerShell Version and Az Module. Check if CCS Agent is already installed on the VM.
- v5: remove the map drive to copy binaries on the VM. It uses UNC path instead.
- v4: add msiexec check on the VM to start the CCS setup only if msiexec is not running
- v3: add CCS deployment information in log file
- v2: use netsh to set firewall rules because PowerShell module "netsecurity" does not exist on Windows 2008 server.
- v1: init
#>

# Check prerequisites
$versionMinimum = [Version]'5.1'
if ($versionMinimum -gt $PSVersionTable.PSVersion) {
													Write-Host "This script requires PowerShell $versionMinimum" -BackgroundColor Red
													Break
}
$CheckAzModule=(Get-InstalledModule -Name Az) 2> $null
if ("$CheckAzModule" -eq "") {
								Write-Host "You need to import PowerShell Az Module to use this script" -BackgroundColor Red
								Break
}

# login to Azure
Login-AzAccount
$global:SubscriptionNameInit = (Get-AzContext).Subscription.Name

# Set the Path for Log File:
if(!(Test-Path -Path "C:\Campbell\")) {New-Item -Path "C:\" -Name "Campbell" -ItemType "directory"}
if(!(Test-Path -Path "C:\Campbell\Logs\")) {New-Item -Path "C:\Campbell\" -Name "Logs" -ItemType "directory"}
$global:txtpath = "C:\Campbell\Logs\"

function Import {
Write-Host "Please select CSV source File of Windows servers" ; ""
Wait-Event -Timeout 2
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files cannot be chosen
	Filter = 'Csv (*.csv)|*.csv' # Specified file types
}
[void]$FileBrowser.ShowDialog()
$global:WindowsFile = $FileBrowser.FileName;
If($FileBrowser.FileNames -like "*\*") {
Write-Host "Windows VM list: $global:WindowsFile" ; ""
Write-Output "********************************************" ; "**         CHECK IMPORT CSV FILE          **" ; "********************************************" ; ""
$global:vms = Import-Csv "$global:WindowsFile" -Delimiter ';'
if($? -eq $True)
{
    # Check headers from CSV File
    $headers = $global:vms[0].psobject.Properties.Name
    $headersrequired = 'Name','ResourceGroupName','SubscriptionName'
    $counter = 3
    Switch($headers)
        {
        {$headersrequired -ccontains $_} {$counter--}
        }
        if($counter -eq 0)
        { Write-Host "File is ok" } else {
											Write-Host "Please correct headers" ; "Should contain 'Name' 'ResourceGroupName' 'SubscriptionName'" ; "" ; "Restart the script" ; ""
											Break
											}
} else {
        Write-Host "Import CSV error" -BackgroundColor Red
        Write-Host "Please Check your Import CSV File" -BackgroundColor Red ; ""
        Break
		}
} else {
		Write-Host "CANCELLED By USER" -foreground Red
		Break
		}
}

function SelectSub {
Write-Output "********************************************" ; "**         SELECTING SUBSCRIPTION         **" ; "********************************************" ; ""
# Show Subscription List to let user select one
Write-Output "List of Subscriptions - Select one from the GUI" ; "-----------------------------------------------" ; ""
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Subscription'
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(300,235)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(380,235)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(480,20)
$label.Text = 'Please select a subcription:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(460,220)
$listBox.Height = 195

$subscription = Get-AzSubscription | where {$_.Status -eq "Enabled"} #{$_.Name -notmatch "AU " -and ($_.Name -notmatch " 1")}
$Listsub = ($subscription.Name | Sort-Object)
if($Listsub.Count -eq 1) {
	[void] $listBox.Items.Add($Listsub)
	} else {
			for ($i=0; $i -le $Listsub.Length - 1 ; $i++)
				{
				$subname = $Listsub[$i]
				[void] $listBox.Items.Add($subname)
				}
			}
$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
$global:SubscriptionName = $listBox.SelectedItem

Write-Host "*******************************************************"
if($global:SubscriptionName -cne $global:SubscriptionNameInit)
                                {
                                Set-AzContext -SubscriptionName "$global:SubscriptionName"
                                $global:SubscriptionNameInit = $global:SubscriptionName
                                }
Write-Host "" ; "Retrieving All Windows VM from the subscription" ; ""
$global:vms = (Get-AzVM -Status | Where-Object {$_.Tags.Keys -match "SupportedBy" -and $_.Tags.Values -match "CloudOPS" -and $_.StorageProfile.OsDisk.OsType -match "Windows" -and $_.PowerState -match "running"})
} else {
		Write-Host "CANCELLED By USER" -foreground Red
		Break
		}
} 

function CheckSub {
                    $DateTime = Get-Date -Format "yyyyMMdd-HHmmss"
                    $txtLogFile = $global:txtpath + "CPBDeployment" + "_" + $global:SubscriptionName + "_" + $DateTime + ".txt"
                    clear
                    Write-Host "OUTPUT LOG File: $txtLogFile" ; "" ; "Work in progress..." ; ""
                    foreach ($a in $global:vms) {
						$subname = $a.SubscriptionName
						$checkpwd = Invoke-AzVMRunCommand -ResourceGroupName $a.ResourceGroupName -VMName $a.Name -CommandId 'RunPowerShellScript' -ScriptPath './cpbdeploy.ps1'
						$a.Name
                        $checkpwd.Value.Message
                        $txt = $checkpwd.Value.Message
						$a.Name | Out-File $txtLogFile -Append -Encoding ascii
						$txt | Out-File $txtLogFile -Append -Encoding ascii
                        $txt = ""
					}
}

function CheckImport {
                        $DateTime = Get-Date -Format "yyyyMMdd-HHmmss"
						$LogFilename = Split-Path $global:WindowsFile -leaf
						$txtLogFile = $global:txtpath + "CPBDeployment" + "_" + $LogFilename + "_" + $DateTime + ".txt"
						clear
						Write-Host "OUTPUT LOG File: $txtLogFile" ; "" ; "Work in progress..." ; ""
						foreach ($a in $vms) {
						    $subname = $a.SubscriptionName
							if($subname -cne $global:SubscriptionNameInit)
                                {
                                Set-AzContext -SubscriptionName "$subname"
                                $global:SubscriptionNameInit = $subname
                                }
							$checkpwd = Invoke-AzVMRunCommand -ResourceGroupName $a.ResourceGroupName -VMName $a.Name -CommandId 'RunPowerShellScript' -ScriptPath './cpbdeploy.ps1'
						    $a.Name
                            $checkpwd.Value.Message
                            $txt = $checkpwd.Value.Message
						    $a.Name | Out-File $txtLogFile -Append -Encoding ascii
						    $txt | Out-File $txtLogFile -Append -Encoding ascii
                            $txt = ""
						}
}

function CreatePS1 {
$cpbdeploy = @'
if (!(Test-Path -Path "C:\Program Files (x86)\Flexera Software\Agent")) {
$UNCPath="\\amazrusepjump02\Share\CPBWinAgent"
# Copy source to local computer
Copy-Item $UNCPath -Destination ".\" -Recurse -Exclude "*.log"
# Create cmd to add Firewall rules using netsh command
#$firewallrule = @"
#netsh advfirewall firewall add rule name=Flexera_Agent_Inbound dir=in action=allow remoteip=192.168.2.10,192.168.2.7 protocol=TCP enable=yes localport=5599-5601 
#netsh advfirewall firewall add rule name=Flexera_Agent_Outbound dir=out action=allow remoteip=192.168.2.10,192.168.2.7 protocol=TCP enable=yes localport=5600 
#"@ | Out-File .\firewall.cmd -Encoding ascii
#start-process -verb runAs ".\firewall.cmd"
# Install from local path
cd .\CPBWinAgent
$FQDN="$env:COMPUTERNAME"
$proc=Get-Process -Name "msiexec*"
if("$proc" -eq "") {
start-process -wait -verb runAs ".\cpbdeploy.cmd"
Write-Host "" ; "Checking Flexera Agent Service..."
Get-Service -ComputerName localhost -Name "ndinit"
} else {Write-Host "Process msiexec is already in used. Wait for process to stop or kill it before starting script again"}
} else {Write-Host "Flexera Agent is already installed on $FQDN"}
'@ | Out-File cpbdeploy.ps1 -Encoding ascii
}

function ClearPS1 {
if (Test-Path .\cpbdeploy.ps1) {Remove-Item .\cpbdeploy.ps1}
}

# Main Menu
Write-Host "Do you want to deploy Flexera agent on Full Subscription or from an input CSV File?" ; "" ; "1: Full Subscription" ; "2: Input CSV File" ; ""
$Selectenv = Read-Host -prompt 'Your choice'
clear
if ($Selectenv -eq "1") {
						SelectSub
                        CreatePS1
						CheckSub
                        ClearPS1
                        }
elseif ($Selectenv -eq "2") {
                            Import
                            CreatePS1
							CheckImport
                            ClearPS1
                            }
else {
		Write-Host ""
        Write-Host "Entry not valid, restart script" -Foreground Red
		}