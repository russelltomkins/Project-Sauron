<#
  .SYNOPSIS
  Name: Create-CustomViews.ps1
  Version: 1.1
  Author: Russell Tomkins - Microsoft Premier Field Engineer
  Blog: https://aka.ms/russellt

  Creates Event Viewer custom views from an input CSV 
  Source: https://www.github.com/russelltomkins/ProjectSauron

  .DESCRIPTION
  Leverages an input CSV file to create custom event views using the xPath 
  filters provided. Can be used to validate xPath filters prior to creating
  input file before creating dedicated custom event channels or for creating
  a friendly customised tree view of events.

  Refer to this blog series for more details
  http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

  .EXAMPLE
  Create the custom views
  Create-CustomViews.ps1 -InputFile DCEvents.csv 

  .PARAMETER InputFile
  A CSV file which must include a ChannelName, ChannelSymbol, QueryPath and the xPath Query itself  
    
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
[CmdletBinding()]
    Param (
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$InputFile)

# Import our Custom Events
$CustomEvents = Import-CSV $InputFile

# We don't care about the providers, just loop through each entry to create the view
ForEach($Channel in $CustomEvents){	

	# Prepare the Channel Details
	$CustomViewName = $Channel.ChannelName.Split("/")[1]
		
	# Convert our ChannelName to create the Subfolder Structure
	$CustomViewNamePath = (($Channel.ChannelName.Split("/"))[0]).replace("-","\")

	# Pre-pend the current Folder path and create the SubFolders
	$ProgramDataPath =  [System.Environment]::ExpandEnvironmentVariables("%Programdata%")
	$CustomViewNamePath = "$ProgramDataPath\Microsoft\Event Viewer\Views\" + $CustomViewNamePath
	New-Item -Type Directory $CustomViewNamePath -Force | out-null
	
	# Create our new XML File
	$xmlFilePath = $CustomViewNamePath + "\" + $Channel.ChannelSymbol + ".xml"
	$XmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$null)

    # Set The Formatting
	$xmlWriter.Formatting = "Indented"
	$xmlWriter.Indentation = "4"

	# Write the XML Decleration
	$xmlWriter.WriteStartDocument()
	
	# Create Instrumentation Manifest
	$xmlWriter.WriteStartElement("ViewerConfig")
		$xmlWriter.WriteStartElement("QueryConfig")
			$xmlWriter.WriteStartElement("QueryParams")
				$xmlWriter.WriteStartElement("UserQuery")
				$xmlWriter.WriteEndElement() # Closing UserQuery
			$xmlWriter.WriteEndElement() # Closing QueryParams
			$xmlWriter.WriteStartElement("QueryNode")
				$xmlWriter.WriteStartElement("Name")
					$xmlWriter.WriteAttributeString("LanguageNeutralValue",$CustomViewName)
				$xmlWriter.WriteEndElement() # Closing Name
				$xmlWriter.WriteStartElement("QueryList")
					$xmlWriter.WriteStartElement("Query")
						$xmlWriter.WriteAttributeString("Id","0")
						$xmlWriter.WriteAttributeString("Path",$Channel.QueryPath)
						$xmlWriter.WriteRaw($Channel.Query)
					$xmlWriter.WriteEndElement() # Closing Query
				$xmlWriter.WriteEndElement() # Closing QueryList
			$xmlWriter.WriteEndElement() # Closing QueryNode
		$xmlWriter.WriteEndElement() # Closing QueryConfig
	$xmlWriter.WriteEndElement() # Closing ViewerConfig
	
	# Close the XML portion of the document
	$xmlWriter.WriteEndDocument()
    
	# Save and close the .XML file
	$xmlWriter.Finalize
	$xmlWriter.Flush()
	$xmlWriter.Close()
}

Write-Host "`nCustom views stored at `"$ProgramDataPath\Microsoft\Event Viewer\Views`""
Write-Host "`nLaunch Event Viwer (eventvwr.exe) and expand Custom Views to use them`n"
# -----------------------------------------------------------------------------------
# End of Script
# -----------------------------------------------------------------------------------
