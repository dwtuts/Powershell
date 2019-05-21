# Script by Micaiah Martin
# Your Zabbix server IP
$ZabbixServerIP = "192.168.1.100"
# Listening port on the Zabbix server
$ListenPort = 10050
$EnableRemoteCommands = 1
$InstallFlags = "qn"
# Grabs hostname from the computer and converts to lowercase so that Zabbix can talk to it nicely
$Hostname = $($env:computername+"."+$env:USERDNSDOMAIN).ToLower()


# Downloads the agent from the web
$zabbixURL = "https://www.zabbix.com/downloads/4.2.1/zabbix_agent-4.2.1-win-amd64-openssl.msi"
$saveLocation = "$PSScriptRoot\zabbix.msi"
$start_time = Get-Date

Invoke-WebRequest -Uri $zabbixURL -OutFile $saveLocation
Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"



# executes the MSI agent
msiexec /I $saveLocation HOSTNAME=$Hostname HOSTNAMEFQDN=1 SERVER=$ZabbixServerIP LPORT=$ListenPort SERVERACTIVE=$ZabbixServerIP /qn

# Opens the port on the Windows Firewall

# Opens Inbound
New-NetFirewallRule -DisplayName "Allow Zabbix Agent Inbound" -Direction Inbound -Program "C:\Program Files\Zabbix Agent\zabbix_agentd.exe" -RemoteAddress LocalSubnet -Action Allow -Profile Any

# Opens Outbound
New-NetFirewallRule -DisplayName "Allow Zabbix Agent Outbound" -Direction Outbound -Program "C:\Program Files\Zabbix Agent\zabbix_agentd.exe" -RemoteAddress LocalSubnet -Action Allow -Profile Any

#Cleanup 
 Remove-Item –path $saveLocation

 #Notification of completion
Add-Type -AssemblyName System.Windows.Forms 
$global:balloon = New-Object System.Windows.Forms.NotifyIcon
$path = (Get-Process -id $pid).Path
$balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
$balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
$balloon.BalloonTipText = 'Zabbix Agent has been installed successfully'
$balloon.BalloonTipTitle = "WDS Monitoring Installed" 
$balloon.Visible = $true 
$balloon.ShowBalloonTip(5000)