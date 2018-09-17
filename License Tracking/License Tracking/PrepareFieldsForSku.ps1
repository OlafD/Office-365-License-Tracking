param (
	[Parameter(Mandatory=$true)]
	[string]$Listname,
	[Parameter(Mandatory=$true)]
	[string]$TranscriptPath,
	[Parameter(Mandatory=$true)]
	$Credentials
)

#------- Constants -------

$LICENSE_TRACKING_LIST = $Listname


#------- Functions -------

function IsMsolServiceConnected
{
    $values = Get-MsolAccountSku -ErrorAction SilentlyContinue

	return ($values -ne $null)
}

function TestSPOConnection
{
	$result = $false
	
	Try
	{
		$ctx = Get-PnPContext

		$result = $true
	}
	Catch
	{
		$result = $false
	}

	Return $result
}

function LoadXmlDocument
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Filename
	)

	$xmlDoc = New-Object System.Xml.XmlDocument
	$xmlDoc.Load($Filename)

	return $xmlDoc
}

function FindLastFieldNumber
{
	param (
		[string]$fieldMappingsFile = "$PSScriptRoot\FieldMappings.xml"
	)

	$xmlDoc = LoadXmlDocument -Filename $fieldMappingsFile
	$xmlRoot = $xmlDoc.DocumentElement

	$result = $xmlRoot.SelectSingleNode("//FieldMappings/Field/@FieldNumber[not(. <=../preceding-sibling::Field/@FieldNumber) and not(. <=../following-sibling::Field/@FieldNumber)]")

	if ($result -eq $null)
	{
		return 0
	}

    $value = $result.'#text'
	return [int]$value
}

function IsSkuInFieldMappings
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[string]$fieldMappingsFile = "$PSScriptRoot\FieldMappings.xml"
	)

	$result = $false

	$xmlDoc = LoadXmlDocument -Filename $fieldMappingsFile
	$xmlRoot = $xmlDoc.DocumentElement

	$xpath = "//FieldMappings/Field[@SkuPartNumber='" + $Sku + "']"
	$nodes = $xmlRoot.SelectNodes($xpath)

	$result = ($nodes.Count -ne 0)

	return $result
}

function AddNewFieldToMappings
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[Parameter(Mandatory=$true)]
		[string]$DisplayName,
		[Parameter(Mandatory=$true)]
		[string]$InternalName,
		[Parameter(Mandatory=$true)]
		[string]$Id,
		[Parameter(Mandatory=$true)]
		[int]$FieldNumber,
		[string]$fieldMappingsFile = "$PSScriptRoot\FieldMappings.xml"
	)

	$xmlDoc = LoadXmlDocument -Filename $fieldMappingsFile

	$newField = $xmlDoc.CreateElement("Field")
	$newField.SetAttribute("DisplayName", $DisplayName)
	$newField.SetAttribute("InternalName", $InternalName)
	$newField.SetAttribute("Id", $Id)
	$newField.SetAttribute("SkuPartNumber", $Sku)
	$newField.SetAttribute("FieldNumber", $FieldNumber.ToString())

	$ignore = $xmlDoc.FieldMappings.AppendChild($newField)
	$xmlDoc.Save($fieldMappingsFile)
}

function AddSkuToThresholds
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[string]$ThresholdsFile = "$PSScriptRoot\SkuThresholds.xml"
	)

	$xmlDoc = LoadXmlDocument -Filename $ThresholdsFile

	$defaultReceipient = $xmlDoc.SkuThresholds.DefaultReceipient

	$newSku = $xmlDoc.CreateElement("Sku")
	$newSku.SetAttribute("Name", $Sku)
	$newSku.SetAttribute("FriendlyName", $Sku)
	$newSku.SetAttribute("Threshold", "0")
	$newSku.SetAttribute("Receipient", $defaultReceipient)

	$ignore = $xmlDoc.SkuThresholds.AppendChild($newSku)
	$xmlDoc.Save($ThresholdsFile)
}

function AddNewFieldToList
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Id,
		[Parameter(Mandatory=$true)]
		[string]$DisplayName,
		[Parameter(Mandatory=$true)]
		[string]$InternalName,
		[switch]$AddToDefaultView
	)

	$xmlDoc = LoadXmlDocument -Filename "$PSScriptRoot\FieldTemplate.xml"
	$xmlDoc.Field.Id = $Id
	$xmlDoc.Field.Name = $InternalName
	$xmlDoc.Field.StaticName = $InternalName
	$xmlDoc.Field.DisplayName = $DisplayName

	$fieldXml = $xmlDoc.InnerXml.ToString()

	$f = Add-PnPFieldFromXml -List $LICENSE_TRACKING_LIST -FieldXml $fieldXml 

	if ($f -ne $null)
	{
		if ($AddToDefaultView.ToBool() -eq $true)
		{
			AddFieldToDefaultView -InternalName $InternalName
		}
	}
}

function GetDefaultReceipient
{
	param (
		[string]$ThresholdsFile = "$PSScriptRoot\SkuThresholds.xml"
	)

	$xmlDoc = LoadXmlDocument -Filename $ThresholdsFile

	$defaultReceipient = $xmlDoc.SkuThresholds.DefaultReceipient

	return $defaultReceipient
}

function AddFieldToDefaultView
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$InternalName
	)

	$view = Get-PnPView -List $LICENSE_TRACKING_LIST -Includes ViewFields
	$view.ViewFields.Add($InternalName)
	$view.Update()
	Invoke-PnPQuery
}

function CreateInternalName
{
	param (
		[Parameter(Mandatory=$true)]
		[int]$Number		
	)

	return "Field" + $Number.ToString("000")
}

function GetReceipientForSku
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[string]$ThresholdFile = "$PSScriptRoot\SkuThresholds.xml"
	)

	$result = ""

	$xmlDoc = LoadXmlDocument -Filename $ThresholdFile
	$xmlRoot = $xmlDoc.DocumentElement

	$xpath = "//SkuThresholds/Sku[@Name='$Sku']"
	$node = $xmlRoot.SelectSingleNode($xpath)

	if ($node -eq $null)
	{
		Write-Host -ForegroundColor Red "Sku $Sku not found in threshold file $ThresholdFile"
	}
	else
	{
		$result = $node.Receipient
	}

	return $result
}

#------- Main -------

Import-Module MSOnline

if ((IsMsolServiceConnected) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a Msol Service. Processing stopped."
	Exit
}

if ((TestSPOConnection) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a SharePoint Online site collection. Processing stopped."
	Exit
}

$transcriptExtension = Get-Date -Format yyyyMMdd-HHmmss
$transcriptFile = "$TranscriptPath\SAP_PrepareFieldsForSku_Transcript_$transcriptExtension.txt"

Start-Transcript -Path $transcriptFile

$x = FindLastFieldNumber 
$nextFieldNumber = $x + 1

$newSkuNotification = @{}

$accountSkuCollection = Get-MsolAccountSku

Write-Host "Found" $accountSkuCollection.Count "items for account sku"

foreach ($accountSku in $accountSkuCollection)
{
	$skuPartNumber = $accountSku.SkuPartNumber

	Write-Host "Processing sku $skuPartNumber"

	if ((IsSkuInFieldMappings -Sku $skuPartNumber) -eq $false)
	{
		Write-Host -ForegroundColor Yellow "- not found in field mappings"

		# Available field

		# field in SharePoint list
		$internalName = CreateInternalName -Number $nextFieldNumber
		$displayName = "$skuPartNumber Available"
		$id = "{" + ([Guid]::NewGuid()).Guid + "}"
		AddNewFieldToList -Id $id -DisplayName $displayName -InternalName $internalName -AddToDefaultView

		Write-Host "- added $displayName to list"

		# FieldMapping entry
		AddNewFieldToMappings -Sku $skuPartNumber -DisplayName $displayName -InternalName $internalName -Id $id -FieldNumber $nextFieldNumber 
		$nextFieldNumber++

		Write-Host "- added $displayName to FieldMapping file"

		# Current field

		# field in SharePoint list
		$internalName = CreateInternalName -Number $nextFieldNumber
		$displayName = "$skuPartNumber Current"
		$id = "{" + ([Guid]::NewGuid()).Guid + "}"
		AddNewFieldToList -Id $id -DisplayName $displayName -InternalName $internalName -AddToDefaultView

		Write-Host "- added $displayName to list"

		# FieldMapping entry
		AddNewFieldToMappings -Sku $skuPartNumber -DisplayName $displayName -InternalName $internalName -Id $id -FieldNumber $nextFieldNumber 
		$nextFieldNumber++

		Write-Host "- added $displayName to FieldMapping file"

		# add sku to Threshold file
		AddSkuToThresholds -Sku $skuPartNumber

		$newSkuNotification.Add($skuPartNumber, $skuPartNumber)

		Write-Host "- added $skuPartNumber to thresholds file"
	}
    else
    {
		Write-Host -ForegroundColor Green "- found in field mappings, nothing to do"
    }
}

if ($newSkuNotification.Count -gt 0)
{
	$receipient = GetDefaultReceipient

	. .\SendMail.ps1 -MailType NewSku -SkuToNotify $newSkuNotification -Receipient $receipient -Credentials $Credentials
}

Write-Host -ForegroundColor Green "Done."

Stop-Transcript
