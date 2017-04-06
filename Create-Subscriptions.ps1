<#
  .SYNOPSIS
  Name: Create-Subscriptions.ps1
  Version: 1.1
  Author: Russell Tomkins - Microsoft Premier Field Engineer
  Blog: https://aka.ms/russellt

  Bulk creation of Windows Event Collection Subscriptions from an input CSV
  Source: https://www.github.com/russelltomkins/ProjectSauron

  .DESCRIPTION
  Leverages an input CSV file to bulk create WEC subscriptions for event delivery
  to dedicated custom event channels. Subscriptions are imported by disabled by default.
  Use the -NoImport and -CreateEnabled switches to override the behaviour.

  Refer to this blog series for more details
  http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

  .EXAMPLE
  Create and Import the WEC subscriptions (disabled by default)
  Create-Subscriptions.ps1 -InputFile DCEvents.csv 
  
  .EXAMPLE
  Create, Import and force enable the WEC subscriptions
  Create-Subscriptions.ps1 -InputFile <inputfile.csv> -CreateEnabled

  .EXAMPLE
  Only create the WEC subscription files, do not import them.
  Create-Subscriptions.ps1 -InputFile <inputfile.csv> -NoImport

  .PARAMETER InputFile
  A CSV file which must include a ChannelName, ChannelSymbol, QueryPath and the xPath Query itself  
  
  .PARAMETER OutputFolder
  The location of the output subscription .xml files. Defaults to "Subscriptions" under the current folder
  
  .PARAMETER CreateEnabled
  Creates and imports the subscriptions but enables them immediately.

  .PARAMETER NoImport
  Creates the subscriptions files, but does not import them

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
    [Parameter(Mandatory=$false)][string]$OutputFolder=$PWD,
    [Parameter(Mandatory=$false)][Switch]$CreateEnabled,
    [Parameter(Mandatory=$false)][Switch]$NoImport)

# Import our Custom Events
$CustomChannels = Import-CSV $InputFile

# Loop through Chanel in input events.
ForEach($Channel in $CustomChannels){	

	# --- Setup the Event Channels ---
	# Bind to the Event Channel
	$EventChannel = Get-WinEvent -ListLog $Channel.ChannelName

	# Do not proceed if we are importing and the logs are still disabled.
	If(!($NoImport)) {
		If (!($EventChannel.IsEnabled)) {
			Write-Host "Error: Event Channel is not Enabled" -Foregroundcolor "Red" -BackGroundColor "Black"
			Write-host "Execute `"Prepare-EventChannels.ps1`" to configure them prior to creating event subscriptions"-Foregroundcolor "Red" -BackGroundColor "Black"
		Exit
		}
	}
	
	# --- Create the Subscription XML's
	# Pre-pend the current Folder path and create the SubFolders
	$SubscriptionNamePath = $OutputFolder + "\Subscriptions"
	If(!(Test-Path $SubscriptionNamePath)){New-Item -Type Directory $SubscriptionNamePath | Out-Null}
	
	# Create our new XML File	
	$xmlFilePath = $SubscriptionNamePath + "\" + $Channel.ChannelSymbol + ".xml"
	$XmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$null)

	# Set The Formatting
	$xmlWriter.Formatting = "Indented"
	$xmlWriter.Indentation = "4"

	# Write the XML Decleration
	$xmlWriter.WriteStartDocument()

	# Create Subscription
	$xmlWriter.WriteStartElement("Subscription")
	$xmlWriter.WriteAttributeString("xmlns","http://schemas.microsoft.com/2006/03/windows/events/subscription")

	$xmlWriter.WriteElementString("SubscriptionId",$Channel.ChannelSymbol)
	$xmlWriter.WriteElementString("SubscriptionType","SourceInitiated")
	$xmlWriter.WriteElementString("Description",$Channel.ChannelName)
	If($CreateEnabled){
		$xmlWriter.WriteElementString("Enabled","true")
	}
	Else{ 
		$xmlWriter.WriteElementString("Enabled","false")
	}
	$xmlWriter.WriteElementString("Uri","http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog")
	$xmlWriter.WriteElementString("ConfigurationMode","Custom")
		$xmlWriter.WriteStartElement("Delivery")
		$xmlWriter.WriteAttributeString("Mode","Push")
			$xmlWriter.WriteStartElement("Batching")
				$xmlWriter.WriteElementString("MaxLatencyTime","30000")
			$xmlWriter.WriteEndElement()	# Close Batching
			$xmlWriter.WriteStartElement("PushSettings")
				$xmlWriter.WriteStartElement("Heartbeat")
					$xmlWriter.WriteAttributeString("Interval","3600000")
				$xmlWriter.WriteEndElement() # Closing Heartbeat
			$xmlWriter.WriteEndElement() # Closing PushSettings
		$xmlWriter.WriteEndElement() # Closing Delivery

		$xmlWriter.WriteStartElement("Query")
	        	$xmlWriter.WriteCData('<QueryList><Query Id="0" Path="' + $Channel.QueryPath + '">' + $Channel.Query + '</Query></QueryList>')
		$xmlWriter.WriteEndElement() # Closing Query
	
		$xmlWriter.WriteElementString("ReadExistingEvents","True")
		$xmlWriter.WriteElementString("TransportName","HTTP")
		$xmlWriter.WriteElementString("ContentFormat","events")
		$xmlWriter.WriteStartElement("locale")	
			$xmlWriter.WriteAttributeString("language","en-US")
		$xmlWriter.WriteEndElement() #Closing Locale

		$xmlWriter.WriteElementString("LogFile",$Channel.ChannelName)
		$xmlWriter.WriteElementString("PublisherName","")
		$xmlWriter.WriteElementString("AllowedSourceNonDomainComputers","")
		
		Switch ($Channel.TargetGroup){
                "Domain Controllers" {$xmlWriter.WriteElementString("AllowedSourceDomainComputers","O:NSG:BAD:P(A;;GA;;;DD)S:")}
                "Domain Computers" {$xmlWriter.WriteElementString("AllowedSourceDomainComputers","O:NSG:BAD:P(A;;GA;;;DC)S:")}
				Default{$xmlWriter.WriteElementString("AllowedSourceDomainComputers","O:NSG:BAD:P(A;;GA;;;"+$Channel.TargetGroup+")S:")}
		}
	$xmlWriter.WriteEndElement()   # Closing Subscription

	# End the XML Document
	$xmlWriter.WriteEndDocument()

	# Finish The Document
	$xmlWriter.Finalize
	$xmlWriter.Flush()
	$xmlWriter.Close()

	# Import the subscription to the server
	If(!($NoImport)){
		# Import the subscription to the server
		$command = "C:\Windows\System32\wecutil.exe"
		$action = "create-subscription"
		& $command $action $xmlfilepath
	}
}

# If we didn't import, write out how to import manually
If($NoImport){
	write-Host "Event Channels updated with required settings"
	write-Host "Subscription files located at $SubscriptionNamePath"
	write-host "Import with wecutil.exe create-subscription <subscription-name>.xml"}

# -----------------------------------------------------------------------------------
# End of Script
# -----------------------------------------------------------------------------------