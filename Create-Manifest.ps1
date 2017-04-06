<#
  .SYNOPSIS
  Name: Create-Manifest.ps1
  Version: 1.1
  Author: Russell Tomkins - Microsoft Premier Field Engineer
  Blog: https://aka.ms/russellt

  Creates a custom event channel manifest file from input CSV 
  Source: https://www.github.com/russelltomkins/ProjectSauron

  .DESCRIPTION
  Leverages an input CSV file to create the required Manifest file for .dll compilation
  Once compiled, can be loaded into a Windows Event Collector to allow custom forwarding
  and long term storage of events.

  Refer to this blog series for more details
  http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

  .EXAMPLE
  Creates the Manfifest file to compile
  Create-Manifest.ps1 -InputFile DCEvents.csv 

  .EXAMPLE
  Creates the Manfifest file to compile along with where the DLL will be located on the WEC server
  Create-Manifest.ps1 -InputFile DCEvents.csv -DLLPath "C:\CustomDLLPath" 

  .PARAMETER InputFile
  A CSV file which must include a ProviderSymbol,ProviderName and ProviderGuid  
   
  .PARAMETER DLLPath
  The folder path where the .dll containing the custom event channels that Windows will load  

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
    [Parameter(Mandatory=$false)][String]$DLLPath="C:\Windows\System32")

# Preparation
$BaseName = (Get-Item $InputFile).BaseName
$BasePathName = "$PWD\$BaseName"

$CustomEventsDLL = $DLLPath + "\" + $BaseName + ".dll"	# The Resource and Message DLL that will be referenced in the manifest.
$CustomEventsMAN = "$BasePathName.man"				# The Manifest file

# Import the events from the input file and extract the Provider details.
$CustomEvents = Import-CSV $InputFile
$Providers = $CustomEvents | Select-Object -Property ProviderSymbol,ProviderName,ProviderGuid -Unique # Extract the provider information from input

# Create The Manifest XML Document
$XmlWriter = New-Object System.XMl.XmlTextWriter($CustomEventsMAN,$null)

# Set The Formatting
$xmlWriter.Formatting = "Indented"
$xmlWriter.Indentation = "4"
 
# Write the XML Decleration
$xmlWriter.WriteStartDocument()

# Create Instrumentation Manifest
$xmlWriter.WriteStartElement("instrumentationManifest")
$xmlWriter.WriteAttributeString("xsi:schemaLocation","http://schemas.microsoft.com/win/2004/08/events eventman.xsd")
$xmlWriter.WriteAttributeString("xmlns","http://schemas.microsoft.com/win/2004/08/events")
$xmlWriter.WriteAttributeString("xmlns:win","http://manifests.microsoft.com/win/2004/08/windows/events")
$xmlWriter.WriteAttributeString("xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance")
$xmlWriter.WriteAttributeString("xmlns:xs","http://www.w3.org/2001/XMLSchema")
$xmlWriter.WriteAttributeString("xmlns:trace","http://schemas.microsoft.com/win/2004/08/events/trace")

# Create Instrumentation, Events and Provider Elements
$xmlWriter.WriteStartElement("instrumentation")
	$xmlWriter.WriteStartElement("events")
	
	$Providers = $CustomEvents | Select-Object -Property ProviderSymbol,ProviderName,ProviderGuid -Unique
	ForEach($Provider in $Providers){
		$xmlWriter.WriteStartElement("provider")
			$xmlWriter.WriteAttributeString("name",($Provider.ProviderName))
			$xmlWriter.WriteAttributeString("guid",$Provider.ProviderGUID)
			$xmlWriter.WriteAttributeString("symbol",$Provider.ProviderSymbol)
			$xmlWriter.WriteAttributeString("resourceFileName",$CustomEventsDLL)
			$xmlWriter.WriteAttributeString("messageFileName",$CustomEventsDLL)
			$xmlWriter.WriteAttributeString("parameterFileName",$CustomEventsDLL)
			$xmlWriter.WriteStartElement("channels")

			$Channels = $CustomEvents | Where-Object{$_.ProviderSymbol -eq $Provider.ProviderSymbol}
			ForEach($Channel in $Channels){	
				$xmlWriter.WriteStartElement("channel")	
					$xmlWriter.WriteAttributeString("name",$Channel.ChannelName)
					$xmlWriter.WriteAttributeString("chid",($Channel.ChannelName).Replace(' ',''))
					$xmlWriter.WriteAttributeString("symbol",$Channel.ChannelSymbol)
					$xmlWriter.WriteAttributeString("type","Admin")
					$xmlWriter.WriteAttributeString("enabled","false")
				$xmlWriter.WriteEndElement() # Closing channel
				}
			$xmlWriter.WriteEndElement() # Closing channels
		$xmlWriter.WriteEndElement() # Closing provider
	}		
	$xmlWriter.WriteEndElement() # Closing events
$xmlWriter.WriteEndElement() # Closing Instrumentation
$xmlWriter.WriteEndElement()   # Closing instrumentationManifest
 
# End the XML Document
$xmlWriter.WriteEndDocument()
 
# Finish The Document
$xmlWriter.Finalize
$xmlWriter.Flush()
$xmlWriter.Close()

# Output the usage instructions
Write-Host "`nThe manifest file has been generated at `"$CustomEventsMAN`"`n"
Write-Host "Step 1: With the Windows 10 SDK installed, open a Command Prompt and change directory to the folder with the .man file (This will not work in PowerShell!) `n"

Write-Host "`t `"C:\Program Files (x86)\Windows Kits\10\bin\x64\mc.exe`" `"$CustomEventsMAN`""
Write-Host "`t `"C:\Program Files (x86)\Windows Kits\10\bin\x64\mc.exe`" -css `"NameSpace`" `"$CustomEventsMAN`""
Write-Host "`t `"C:\Program Files (x86)\Windows Kits\10\bin\x64\rc.exe`" `"$BasePathName.rc`""
Write-Host "`t `"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`" /win32res:`"$BasePathName.res`" /unsafe /target:library /out:`"$BasePathName.dll`" `"$BasePathName.cs`"`n"

Write-Host "Step 2: On the WEC server, copy both the .man and .dll file to $DLLPath"
Write-Host "Step 3: Load the custom event channels by executing:`n"
Write-Host "`t `"c:\windows\system32\wevtutil.exe`" im `"$DLLPath\$BaseName.man`"`n"

# -----------------------------------------------------------------------------------
# Main Script
# -----------------------------------------------------------------------------------