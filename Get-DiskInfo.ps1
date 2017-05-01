﻿#=======================================================================================
# Get-LockedOutUserDetails
# Created on: 2017-01-11
# Version 1.0
# Last Updated: 
# Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@lucky13solutions.com
#
# Purpose: This script pulls the volume information for a remote server using WMI and 
#          returns the info to the user in a Grid View
#
# Notes: 
# 
# Change Log:
# 
#
#=======================================================================================
##
# Clear the screen
#
Clear-Host
#
#
# Load required components to use a Popup Window to get user input using VB
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName Microsoft.VisualBasic
#
# Get Server name from user
#
$ServerName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the name of the server that you would like to get the disk info on','Server Name', 'ServerName')
#
# Hash Table for Drive Type
#
$hash = @{
   "2" = "Removable disk"
   "3" = "Fixed local disk"
   "4" = "Network disk"
   "5" = "Compact disk"
   }
#
#
# Declare Empty Arrays
$ServerDisksInfo = @()
$ServerDiskDetailInfo = @()
# $AllServerDiskInfo = @()
#
$AllServerDiskInfo = Get-WMIObject Win32_Volume -ComputerName $ServerName
#
# $ServerDisksInfo = $AllServerDiskInfo | Where-Object {$_.Size -gt -1}
ForEach ($TempDiskInfo in $AllServerDiskInfo) {
    $Temp = New-Object PSObject
    $TempDriveType = $TempDiskInfo.DriveType.ToString()
    $Temp | Add-Member -MemberType NoteProperty -Name Date -Value (Get-Date)
    $Temp | Add-Member -MemberType NoteProperty -Name Server -Value $ServerName
    $Temp | Add-Member -MemberType NoteProperty -Name DriveLetter -Value $TempDiskInfo.Name
    $Temp | Add-Member -MemberType NoteProperty -Name VolumeName -Value $TempDiskInfo.Label
    $Temp | Add-Member -MemberType NoteProperty -Name DriveType -Value ($Hash.Item($TempDriveType))
    $Temp | Add-Member -MemberType NoteProperty -Name FileSystem -Value $TempDiskInfo.FileSystem
    $Temp | Add-Member -MemberType NoteProperty -Name "Size(GB)" -Value ([Math]::Round($TempDiskInfo.Capacity / 1GB,2))
    $Temp | Add-Member -MemberType NoteProperty -Name "FreeSpace(GB)" -Value ([Math]::Round($TempDiskInfo.FreeSpace / 1GB,2))
    IF($TempDiskInfo.Capacity -gt 0) {$Temp | Add-Member -MemberType NoteProperty -Name "%Free" -Value ([Math]::Round(($TempDiskInfo.FreeSpace/$TempDiskInfo.Capacity)*100,2))} 
    Else {$Temp | Add-Member -MemberType NoteProperty -Name "%Free" -Value 0}
    $ServerDiskDetailInfo += $Temp
    }
$ServerDiskDetailInfo | Out-GridView
[System.Windows.Forms.MessageBox]::Show('Please click OK to close the script','End Script')