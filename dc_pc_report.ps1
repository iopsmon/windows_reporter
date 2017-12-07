#===============================================================
#Author: Deepak Chohan
#Script Name: dc_pc_report.ps1
#Description: All PC Info - HTML Based
#Date: 02/10/16
#Version:
#===============================================================
#Updates
#
#===============================================================
#Notes
#Note1: The Paths maybe different, so change this if need be.
#Note2: Powershell Access maybe disabled, so you may need access.
#Note3: Check the varibles to you local system, they may beed to be changed
#===============================================================
#Powershell Security Notes
#Set-ExecutionPolicy Unrestricted (allows powershell)
#Restricted -- Restricted is the default execution policy and locks PowerShell down so that commands can be entered only interactively. PowerShell scripts are not allowed to run.
#All Signed -- If the execution policy is set to All Signed then scripts will be allowed to run, but only if they are #signed by a trusted publisher.
#Remote Signed -- If the execution policy is set to Remote Signed, any PowerShell scripts that have been locally created will be allowed to run. Scripts created remotely are allowed to run only if they are signed by a trusted publisher.
#Unrestricted -- As the name implies, Unrestricted removes all restrictions from the execution policy.
#===============================================================

#*** My General Variables Section ***
$LComputerName = (Get-Item env:\Computername).Value
$Yesterday = (Get-Date) - (New-TimeSpan -Day 1)
$MyPath = "."

#For IE

$MyUrlFile = ".\$LComputerName.Report.htm"

#*** Event Log varibles Section ***
$Log1 = "Application"
$Log2 = "System"
$Log3 = "Security"

#***HTML variables Section ***
$MyTitle = "Windows Process for $LComputerName"
$MyDate = (get-Date)
#Titles

$MyPre = "<h1>iOpsMon PC Reporter ©</h1>"

#Images

$MyImageLogo = "<h2><center><strong>PC Report: $LComputerName On: $MyDate </strong></center></h2>"

#Processes and Services titles
$MyPT1 = "<h2>Process Table</h2>"
$MySVCT1 = "<h2>Services Table</h2>"

#Event titles
$MyET1 = "<h2>Application Events</h2>"
$MyET2 = "<h2>System Events</h2>"
$MyET3 = "<h2>Security Events</h2>"
$MyPost = "<h2>Development By Deepak Chohan © 2017 </h2>"

#Disk titles
$MyDiskT = "<h2>Disk Report</h2>"

#Software title
$MySoftT = "<h2>Software Inventory</h2>"

#Performance
$MyPerfT = "<h2>System Performance</h2>"

#OS title
$MyOST = "<h2>System OS Info</h2>"

#Network title
$MyNetworkT = "<h2>Network Info</h2>"

#Shares title
$MySharesT = "<h2>Shares OS Info</h2>"

#Firewall title
$MyFWIT = "<h2>Firewall Inbound Status</h2>"

#Shares title
$MyFWOUT = "<h2>Firewall OutBound Status</h2>"

#html tags
$MyBr = "<br>"

#***Data Collection Section ***

#OS Info

$computerSystem = Get-wmiobject Win32_ComputerSystem
$computerName = Get-WmiObject Win32_OperatingSystem
$computerArc = Get-WmiObject Win32_OperatingSystem
$computerSp = Get-WmiObject Win32_OperatingSystem
$computerCPU = Get-wmiobject Win32_Processor
$computerRam = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | Foreach {"{0:N2}" -f ([math]::round(($_.Sum / 1GB),2))}
$computerModel = (Get-wmiobject Win32_ComputerSystem).Model
#Use this to get a list of propertys
#Get-WmiObject Win32_OperatingSystem | Get-Member

$OSObject = New-Object PSObject -property @{
'PCName' = $computerSystem.Name
'OS' = $computerName.caption
'SP' = $computerSp.ServicePackMajorVersion
'Arc' = $computerArc.OSArchitecture
'Model' = $computerModel
'CPU_NumberOfCores' = $computerCPU.NumberOfCores
'CPU_Name' = $computerCPU.Name
'CPU_Description' = $computerCPU.Description
'CPU_Manufacturer' = $computerCPU.Manufacturer
'Manufacturer' = $computerSystem.Manufacturer
'Memory' = $computerRam
} $OSObject | Select PCName, OS, SP, Arc, Model, CPU_NumberOfCores, CPU_Name, CPU_Manufacturer, Memory | ConvertTo-HTML -Fragment

#Export the fields you want from above in the specified order

Start-Sleep -s 1

#Performance Data

$AVGProc = Get-WmiObject win32_processor | Measure-Object -property LoadPercentage -Average | Select Average
$OS = gwmi -Class win32_operatingsystem |
Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}

$vol1 = Get-WmiObject -Class win32_Volume -Filter "DriveLetter = 'C:'" |
Select-object @{Name = "C PercentFree"; Expression = {“{0:N2}” -f (($_.FreeSpace / $_.Capacity)*100) } }

$vol2 = Get-WmiObject -Class win32_Volume -Filter "DriveLetter = 'D:'" |
Select-object @{Name = "D PercentFree"; Expression = {“{0:N2}” -f (($_.FreeSpace / $_.Capacity)*100) } }

$vol3 = Get-WmiObject -Class win32_Volume -Filter "DriveLetter = 'E:'" |
Select-object @{Name = "E PercentFree"; Expression = {“{0:N2}” -f (($_.FreeSpace / $_.Capacity)*100) } }

$MyStatObject = New-Object PSObject -property @{
'PCName' = "$LComputerName"
'CPULoad' = "$($AVGProc.Average)%"
'MemLoad' = "$($OS.MemoryUsage)%"
'C-Drive' = "$($vol1.'C PercentFree')%"
'D-Drive' = "$($vol2.'D PercentFree')%"

} | Select PCName, CPULoad, MemLoad, C-Drive, D-Drive | ConvertTo-HTML -Fragment

start-sleep -s 1

#Process Information

$MyProcess = Get-Process | Sort-Object -Descending WS | select Name, Path, WS, CPU, Company | ConvertTo-HTML -Fragment
Start-Sleep -s 1

#Services
$MyServices = get-WmiObject Win32_service | Where-Object {$_.Startmode -contains 'Auto'} |
select PSComputerName, Name, PathName, ProcessId, StartMode, StartName, State | ConvertTo-HTML -Fragment

#Disk Data
$MyDisk = Get-Volume | Where { $_.DriveType -eq 'Fixed' -Or $_.DriveType -eq 'Removable' } |
Select DriveLetter, Filesystem, DriveType, Size, SizeRemaining | ConvertTo-HTML -Fragment

#Network Info
$MyNetworkInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE |
select DHCPEnabled, @{L='IPAddress';ex={$_.IPAddress}},
@{L='DefaultIPGateway';ex={$_.DefaultIPGateway}},
@{L='ServiceName';ex={$_.ServiceName}},
@{L='Description';ex={$_.Description}},
@{L='DNSDomain';ex={$_.DNSDomain}} | ConvertTo-HTML -Fragment
Start-Sleep -s 1

#Local Shares
$MyShares = Get-WmiObject -Class Win32_Share | select PSComputerName, Name | ConvertTo-HTML -Fragment
Start-Sleep -s 1

#Software Inventory

$MySoftInstall = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | ConvertTo-HTML -Fragment

Start-Sleep -s 1

#All Event Logs from 1 day, date / time taken from when the script is run

$MyEventApp = Get-WinEvent @{LogName=$Log1; Level=2,3; StartTime=$Yesterday} -MaxEvents 100 |
select Id, TimeCreated, Message, LevelDisplayName | ConvertTo-HTML -Fragment
Start-Sleep -s 1

$MyEventSys = Get-WinEvent -FilterHashTable @{LogName=$Log2; Level=2,3; StartTime=$Yesterday} |
select Id,TimeCreated, Message, LevelDisplayName | ConvertTo-HTML -Fragment
Start-Sleep -s 1

$MyEventSec = Get-WinEvent -FilterHashTable @{LogName=$Log3; StartTime=$Yesterday} -MaxEvents 100 |
select Id,TimeCreated, Message, LevelDisplayName | ConvertTo-HTML -Fragment
Start-Sleep -s 1

#Firewall Status
#Inbound Rules
$MyFWInbound = Get-NetFirewallRule | Where { $_.Enabled –eq ‘True’ –and $_.Direction –eq ‘Inbound’ } |
Select DisplayName, Direction, Action | ConvertTo-HTML -Fragment

#OutBound Rules
$MyFWOutbound = Get-NetFirewallRule | Where { $_.Enabled –eq ‘True’ –and $_.Direction –eq ‘Outbound’ } |
Select DisplayName, Direction, Action | ConvertTo-HTML -Fragment

Start-Sleep -s 1

# HTML Body Content, this has all the outputs from the commands as variable

$MyBodyAll1 = "$MyPre $MyImageLogo $MyBr $MyOST $MyBr $OSObject $MyBr $MyPerfT $MyStatObject"

$MyBodyAll2 = "$MyBr $MyPT1 $MyProcess $MyBr $MySVCT1 $MyServices $MyBr $MyDiskT $MyDisk"

$MyBodyAll3 = "$MyBr $MyNetworkT $MyNetworkInfo $MyBr $MySharesT $MyShares"

$MyBodyAll4 = "$MyBr $MySoftT $MySoftInstall"

$MyBodyAll5 = "$MyBr $MyET2 $MyEventSys $MyBr $MyET1 $MyEventApp $MyBr $MyET3 $MyEventSec"

$MyBodyAll6 = "$MyBr $MyFWIT $MyFWInbound $MyBr $MyFWOUT $MyFWOutbound "

#This converts all the variables into one htm file and uses the css file
ConvertTo-Html -Title $MyTitle -Body "$MyBodyAll1 $MyBodyAll2 $MyBodyAll3 $MyBodyAll4 $MyBodyAll5 $MyBodyAll6 " `
-post $MyPost -CSSUri ".\dc_3.css" | out-file "$MyUrlFile"


