<#
  .SYNOPSIS
  Name: Create-Subscriptions.ps1
  Version: 1.0
  Author: Russell Tomkins - Microsoft Premier Field Engineer
  Blog: https://aka.ms/russellt

  Bulk creation of Windows Event Collection Subscriptions from  input CSV
  Source: https://www.github.com/russelltomkins/ProjectSauron

  .DESCRIPTION
  Leverages an input CSV file to bulk create WEC subscriptions for event delivery
  to dedicated custom event channels

  Refer to this blog series for more details
  http://blogs.technet.microsoft.com/russellt/2017/03/23/project-sauron-part-1

  .EXAMPLE
  Create, Import and Enable the WEC subscriptions.
  Create-Subscriptions.ps1 -InputFile DCEvents.csv 
  
  .EXAMPLE
  Create, Import but don't enable the WEC subscriptions
  Create-Subscriptions.ps1 -InputFile <inputfile.csv> -CreateDisabled

  .EXAMPLE
  Only create the WEC subscription files, do not import them.
  Create-Subscriptions.ps1 -InputFile <inputfile.csv> -NoImport

  .PARAMETER InputFile
  A CSV file which must include a ChannelName, ChannelSymbol, QueryPath and the xPath Query itself  
  
  .PARAMETER LogRootPath
  The location of .evtx event log files. Defaults to "D:\Logs"  

  .PARAMETER OutputFile
  The location of the output subscription .xml files. Defaults to "D:\Logs"  
  
  .PARAMETER CreateDisabled
  Creates and imports the subscriptions, but does not enable it

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
    [Parameter(Mandatory=$false)][String]$LogRootPath="D:\Logs",
    [Parameter(Mandatory=$false)][string]$OutputFolder=$PWD,
    [Parameter(Mandatory=$false)][Switch]$CreateDisabled,
	[Parameter(Mandatory=$false)][Switch]$NoImport)

# Import our Custom Events
$CustomChannels = Import-CSV $InputFile

# Create and ACL the Log Roots Folder to allow Network Service access.
If(!(Test-Path $LogRootPath )){New-Item -Type Directory $LogRootPath}
$ACE = New-Object System.Security.AccessControl.FileSystemAccessRule("NETWORK SERVICE",'Modify','ContainerInherit,ObjectInherit','None','Allow')
$LogRootPathACL = (Get-Item $LogRootPath) | Get-ACL
$LogRootPathACL.AddAccessRule($ACE)
$LogRootPathACL | Set-ACL

# Loop through Chanel in input events.
ForEach($Channel in $CustomChannels){	

	# --- Setup the Event Channels ---
	# Bind to the Event Channel
	$EventChannel = Get-WinEvent -ListLog $Channel.ChannelName

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
	
	# --- Create the Subscription XML's
	# Pre-pend the current Folder path and create the SubFolders
	$SubscriptionNamePath = $OutputFolder + "\Subscriptions"
	If(!(Test-Path $SubscriptionNamePath)){New-Item -Type Directory $SubscriptionNamePath}
	
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
	If($CreateDisabled){
		$xmlWriter.WriteElementString("Enabled","false")
	}
	Else{ 
		$xmlWriter.WriteElementString("Enabled","true")
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
# SIG # Begin signature block
# MIIgVAYJKoZIhvcNAQcCoIIgRTCCIEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0Osw/1T4Td6An
# uktM5rKr0UFEp2V+3sHBob/Pz2ZvRKCCG14wggO3MIICn6ADAgECAhAM5+DlF9hG
# /o/lYPwb8DA5MA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBa
# Fw0zMTExMTAwMDAwMDBaMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBAK0OFc7kQ4BcsYfzt2D5cRKlrtwmlIiq9M71IDkoWGAM+IDaqRWVMmE8
# tbEohIqK3J8KDIMXeo+QrIrneVNcMYQq9g+YMjZ2zN7dPKii72r7IfJSYd+fINcf
# 4rHZ/hhk0hJbX/lYGDW8R82hNvlrf9SwOD7BG8OMM9nYLxj+KA+zp4PWw25EwGE1
# lhb+WZyLdm3X8aJLDSv/C3LanmDQjpA1xnhVhyChz+VtCshJfDGYM2wi6YfQMlqi
# uhOCEe05F52ZOnKh5vqk2dUXMXWuhX0irj8BRob2KHnIsdrkVxfEfhwOsLSSplaz
# vbKX7aqn8LfFqD+VFtD/oZbrCF8Yd08CAwEAAaNjMGEwDgYDVR0PAQH/BAQDAgGG
# MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFEXroq/0ksuCMS1Ri6enIZ3zbcgP
# MB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUA
# A4IBAQCiDrzf4u3w43JzemSUv/dyZtgy5EJ1Yq6H6/LV2d5Ws5/MzhQouQ2XYFwS
# TFjk0z2DSUVYlzVpGqhH6lbGeasS2GeBhN9/CTyU5rgmLCC9PbMoifdf/yLil4Qf
# 6WXvh+DfwWdJs13rsgkq6ybteL59PyvztyY1bV+JAbZJW58BBZurPSXBzLZ/wvFv
# hsb6ZGjrgS2U60K3+owe3WLxvlBnt2y98/Efaww2BxZ/N3ypW2168RJGYIPXJwS+
# S86XvsNnKmgR34DnDDNmvxMNFG7zfx9jEB76jRslbWyPpbdhAbHSoyahEHGdreLD
# +cOZUbcrBwjOLuZQsqf6CkUvovDyMIIFLDCCBBSgAwIBAgIQDhlON30mOhkOirPI
# WrUoYzANBgkqhkiG9w0BAQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhE
# aWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE3MDMy
# NzAwMDAwMFoXDTE4MDQwNDEyMDAwMFowaTELMAkGA1UEBhMCQVUxEzARBgNVBAgT
# ClF1ZWVuc2xhbmQxETAPBgNVBAcTCEJyaXNiYW5lMRgwFgYDVQQKEw9SdXNzZWxs
# IFRvbWtpbnMxGDAWBgNVBAMTD1J1c3NlbGwgVG9ta2luczCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAL9yEH4Y+mOkq5qq1yIMMQxZks06om9d6ifoWnQZ
# LwleCoIohbxLcc9RsAsY3b0E0alY/WGBbvxrAXDsfNtV2oRBwq4I1wRbrazuYdec
# V/ON+0cOKvSN3df9AJmbw53MBqlOLJr+f3IyLan40iY2PCt/N12zKVvPnFtoP+Lr
# QwLkUTMT+5LdmGl0UfaLkgno7EG+7CXKL1QDIw1NLiYkw1fxlcu8+MOslqV6ZFVm
# rhrM+Q0tzvVtq4DWSyn63U8j8Ij9cjnPpG3mABFN1dpu31yFBYogcPvFfQzx013f
# s4GI4mu70CDCy1vbi3oSa3jjiqExysDXcOHhZ4RVZ3xKUAsCAwEAAaOCAcUwggHB
# MB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBSiIVol
# K54Mdi8hZEbQ+ZcbWmjObTAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2lj
# ZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZI
# AYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9D
# UFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0Eu
# Y3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAPLir+VRKD+MIfvl
# S7s8KtE6sBOx2JCNewUh4JVtmQECTTpvKvx25TYO23MrApApfhc8qa2mkHNpyjMX
# U7SZog3mNSIJlQrhiF1Y6xNafqbDz31qGU/booX2AHV1yfJbXNWw2tTnbukdhFO/
# 2vSKdUqJZbYp2A+dx5zemxvtf46CTy4PxrcKmn+Umd+Cil3O3TlDTy0LGfzPTL1f
# IOAqtc4bbge6pMn5BwV0dxOZ4JTIsXlFzzIKjjOUNX/+0/iGoYAXvkyOA0wdEiDN
# qug5CTbskpE/ltGa0XCSkglk2j4431JgUC+ew2YgSsEq0dukmdUjz3HpdvrMEYfg
# T5PcXa4wggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEB
# CwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQg
# SUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBD
# b2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD4
# 07Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnU
# NxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSx
# KCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxw
# oqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOs
# Z061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQ
# D2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEA
# MA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcB
# AQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggr
# BgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0
# cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6ch
# nfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToa
# xO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMO
# E0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8
# LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKy
# NookAv4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtP
# xpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggZqMIIF
# UqADAgECAhADAZoCOv9YsWvW1ermF/BmMA0GCSqGSIb3DQEBBQUAMGIxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTAeFw0x
# NDEwMjIwMDAwMDBaFw0yNDEwMjIwMDAwMDBaMEcxCzAJBgNVBAYTAlVTMREwDwYD
# VQQKEwhEaWdpQ2VydDElMCMGA1UEAxMcRGlnaUNlcnQgVGltZXN0YW1wIFJlc3Bv
# bmRlcjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKNkXfx8s+CCNeDg
# 9sYq5kl1O8xu4FOpnx9kWeZ8a39rjJ1V+JLjntVaY1sCSVDZg85vZu7dy4XpX6X5
# 1Id0iEQ7Gcnl9ZGfxhQ5rCTqqEsskYnMXij0ZLZQt/USs3OWCmejvmGfrvP9Enh1
# DqZbFP1FI46GRFV9GIYFjFWHeUhG98oOjafeTl/iqLYtWQJhiGFyGGi5uHzu5uc0
# LzF3gTAfuzYBje8n4/ea8EwxZI3j6/oZh6h+z+yMDDZbesF6uHjHyQYuRhDIjegE
# YNu8c3T6Ttj+qkDxss5wRoPp2kChWTrZFQlXmVYwk/PJYczQCMxr7GJCkawCwO+k
# 8IkRj3cCAwEAAaOCAzUwggMxMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAA
# MBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMIIBvwYDVR0gBIIBtjCCAbIwggGhBglg
# hkgBhv1sBwEwggGSMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5j
# b20vQ1BTMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBm
# ACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABp
# AHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAg
# AEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAg
# AFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAg
# AHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBu
# AGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBp
# AG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTAfBgNV
# HSMEGDAWgBQVABIrE5iymQftHt+ivlcNK2cCzTAdBgNVHQ4EFgQUYVpNJLZJMp1K
# Knkag0v0HonByn0wfQYDVR0fBHYwdDA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwOKA2oDSGMmh0dHA6Ly9j
# cmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMHcGCCsG
# AQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURDQS0xLmNydDANBgkqhkiG9w0BAQUFAAOCAQEAnSV+GzNNsiaB
# XJuGziMgD4CH5Yj//7HUaiwx7ToXGXEXzakbvFoWOQCd42yE5FpA+94GAYw3+pux
# nSR+/iCkV61bt5qwYCbqaVchXTQvH3Gwg5QZBWs1kBCge5fH9j/n4hFBpr1i2fAn
# PTgdKG86Ugnw7HBi02JLsOBzppLA044x2C/jbRcTBu7kA7YUq/OPQ6dxnSHdFMoV
# XZJB2vkPgdGZdA0mxA5/G7X1oPHGdwYoFenYk+VVFvC7Cqsc21xIJ2bIo4sKHOWV
# 2q7ELlmgYd3a822iYemKC23sEhi991VUQAOSK2vCUcIKSK+w1G7g9BQKOhvjjz3K
# r2qNe9zYRDCCBs0wggW1oAMCAQICEAb9+QOWA63qAArrPye7uhswDQYJKoZIhvcN
# AQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJl
# ZCBJRCBSb290IENBMB4XDTA2MTExMDAwMDAwMFoXDTIxMTExMDAwMDAwMFowYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0x
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBzQHDS
# nlZUXKnE0kEGj8kz/E1FkVyBn+0snPgWWd+etSQVwpi5tHdJ3InECtqvy15r7a2w
# cTHrzzpADEZNk+yLejYIA6sMNP4YSYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD/6b3
# +6LNb3Mj/qxWBZDwMiEWicZwiPkFl32jx0PdAug7Pe2xQaPtP77blUjE7h6z8rwM
# K5nQxl0SQoHhg26Ccz8mSxSQrllmCsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2zkBd
# XPao8S+v7Iki8msYZbHBc63X8djPHgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9BwSi
# CQIDAQABo4IDejCCA3YwDgYDVR0PAQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsGAQUF
# BwMBBggrBgEFBQcDAgYIKwYBBQUHAwMGCCsGAQUFBwMEBggrBgEFBQcDCDCCAdIG
# A1UdIASCAckwggHFMIIBtAYKYIZIAYb9bAABBDCCAaQwOgYIKwYBBQUHAgEWLmh0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFk
# BggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBz
# ACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBz
# ACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBD
# AGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBp
# AG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBo
# ACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBl
# ACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAg
# AHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQIMAYB
# Af8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4
# oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJv
# b3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYEFBUAEisTmLKZB+0e36K+Vw0r
# ZwLNMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEB
# BQUAA4IBAQBGUD7Jtygkpzgdtlspr1LPUukxR6tWXHvVDQtBs+/sdR90OPKyXGGi
# nJXDUOSCuSPRujqGcq04eKx1XRcXNHJHhZRW0eu7NoR3zCSl8wQZVann4+erYs37
# iy2QwsDStZS9Xk+xBdIOPRqpFFumhjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+cdkvt
# X8JLFuRLcEwAiR78xXm8TBJX/l/hHrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeSDY2x
# aYxP+1ngIw/Sqq4AfO6cQg7PkdcntxbuD8O9fAqg7iwIVYUiuOsYGk38KiGtSTGD
# R5V3cdyxG0tLHBCcdxTBnU8vWpUIKRAmMYIETDCCBEgCAQEwgYYwcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQQIQDhlON30mOhkOirPIWrUoYzANBglghkgBZQMEAgEFAKCBhDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEi
# BCDk1IB2qVR9RaXlfijXEmFLt+9dHQ5rQkHDcaX4FmgzfjANBgkqhkiG9w0BAQEF
# AASCAQByFSrKaw/KQws3vuIHuFkP8ed1mb/ZKExVBKACbvX8d5XjZXLfQMtWXKtP
# wsRV2vDpsDJRzE5iqjpGNwRTMflRprkwU0MgpFpZd3VzUX+9PlXPUin/H07Ik8Kv
# djn7YzppOMvx7UTeBbMLhMJPJsnaISyffCgeBtEU1zi1I0Fkwy3fUS8Q4A3klQJd
# pWhgUr9esMMr7YQo0z58T4Qhz4EZyLSyrxKhwuxg+belv6/dClgqxdXB9cqge3/2
# J14Pkp2ih2VJy6w+oKfu0G4dp1C/Neh/zzNsjGx5YfwYo1yQKnHnp4YZ9X/oNrwH
# DJooWpB+uwngpFsyd3LKFm1tErhAoYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4
# AgEBMHYwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJl
# ZCBJRCBDQS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNzAzMjgwMTE2NDZa
# MCMGCSqGSIb3DQEJBDEWBBT9yDxcEWrXneHdE8PrXR3ZC/CcrDANBgkqhkiG9w0B
# AQEFAASCAQBMCoNP712CMHL+XJV/OIkJrpashiwLxFPL6KKyggEfcKwRA5k2zNSz
# Mt3B8UiOyl9Qocmxex7T0rwGxRxrcSgFYlKwSngdAKqABTzApFaXzZ6NAhn9eJAd
# zYql9frJD2sAam9My5MhMoGqwbYlKlLlTas1j/maimIZm9/JGgpLqKOBxxKRjF+G
# O+RXU38IZW0DjL64UAKXzB/C9Ybns3R2JYzhwdy5fxGnKb4JLVsV6IiM/oLtAMv9
# Y2FgI9pz0CU6NGsM/eo1thaMNcN3zU2CpcOryiLEHH51t3z5O53aZ5oXHLBG6c5Q
# xGQyvvmL3sBDQcpl/SfhASHvTwlkLdCw
# SIG # End signature block
