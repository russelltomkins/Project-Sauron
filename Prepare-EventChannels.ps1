<#
  .SYNOPSIS
  Name: Prep-EventChannels.ps1
  Version: 1.1
  Author: Russell Tomkins - Microsoft Premier Field Engineer
  Blog: https://aka.ms/russellt

  Preparation of event channels to receive event collection subscriptions from an input CSV
  Source: https://www.github.com/russelltomkins/ProjectSauron

  .DESCRIPTION
  Leverages an input CSV file to prepare the custom event channels created by Create-Manifest.ps1

  Refer to this blog series for more details
  http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

  .EXAMPLE
  Prepare the Event Chanenls using the Input CSV file.
  Create-Subscriptions.ps1 -InputFile DCEvents.csv 
  
  .PARAMETER InputFile
  A CSV file which must include a ChannelName, ChannelSymbol, QueryPath and the xPath Query itself  
  
  .PARAMETER LogRootPath
  The location of .evtx event log files. Defaults to "D:\Logs"  

  LEGAL DISCLAIMER
  This Sample Code is provided for the purpose of illustration only and is not
  intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
  RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
  EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
  MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
  nonexclusive, royalty-free right to use and modify the Sample Code and to
  reproduce and distribute the object code form of the Sample Code, provided
  that You agree: (i) to not use Our name, logo, or trademarks to market Your
  software product in which the Sample Code is embedded; (ii) to include a valid
  copyright notice on Your software product in which the Sample Code is embedded;
  and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
  against any claims or lawsuits, including attorneys fees, that arise or result
  from the use or distribution of the Sample Code.
   
  This posting is provided "AS IS" with no warranties, and confers no rights. Use
  of included script samples are subject to the terms specified
  at http://www.microsoft.com/info/cpyright.htm.
  #>
# -----------------------------------------------------------------------------------
# Main Script
# -----------------------------------------------------------------------------------

# Prepare the Input Paremeters
[CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$InputFile,
    [Parameter(Mandatory=$false)][String]$LogRootPath="D:\Logs")

# Import our Custom Events
$CustomChannels = Import-CSV $InputFile

# Create The Folder
If(!(Test-Path $LogRootPath )){New-Item -Type Directory $LogRootPath | Out-Null}

# Add an ACE to allow LOCAL SERVICE to modify the folder
$ACE = New-Object System.Security.AccessControl.FileSystemAccessRule("LOCAL SERVICE",'Modify','ContainerInherit,ObjectInherit','None','Allow')
$LogRootPathACL = (Get-Item $LogRootPath) | Get-ACL
$LogRootPathACL.AddAccessRule($ACE)
$LogRootPathACL | Set-ACL

# Enable NTFS compression to save disk space
$Query = "select * from CIM_Directory where name = `"$($LogRootPath.Replace('\','\\'))`""
$Results = Invoke-CimMethod -Query $Query -MethodName Compress

# Loop through Chanell form the InputCSV
ForEach($Channel in $CustomChannels){	

	# --- Setup the Event Channels ---
	# Bind to the Event Channel
	$EventChannel = Get-WinEvent -ListLog $Channel.ChannelName -ErrorAction "SilentlyContinue"
	If ($EventChannel -eq $Null){
		Write-Host "`nError: Event channel not loaded:`"$($Channel.ChannelName)`"" -ForeGroundColor Red
		Write-Host "`nEnsure the manifest and dll has been loaded with wevtutil.exe im <path to manifest.man>`n" -foregroundColor Green
	Exit
	}

	# Disable the channel to allow changes
	If ($EventChannel.IsEnabled) {
  		$EventChannel.IsEnabled = $False
		$EventChannel.SaveChanges()
	}
    
	  # Update the channel to our requried Values
	  $NewLogFilePath = $LogRootPath + "\" + $Channel.ChannelSymbol + ".evtx"
	  $EventChannel.LogFilePath = $NewLogFilePath
	  $EventChannel.LogMode = "AutoBackup"
	  $EventChannel.MaximumSizeInBytes = 1073741824
	  $EventChannel.SaveChanges()
      
    # Enable the Log
    $EventChannel.IsEnabled = $True	
    $EventChannel.SaveChanges()
}
# -----------------------------------------------------------------------------------
# End of Script
# -----------------------------------------------------------------------------------