param (
	[Parameter(Mandatory=$true)]
	[string]$TranscriptPath
)

#------- Constants -------

$LICENSE_TRACKING_LIST = "License Tracking"


#------- Functions -------

function IsMsolServiceConnected
{
    $values = Get-MsolAccountSku -ErrorAction SilentlyContinue

	return ($values -ne $null)
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

function GetInternalFieldname
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$DisplayName,
		[string]$fieldMappingsFile = "$PSScriptRoot\FieldMappings.xml"
	)

	$xmlDoc = LoadXmlDocument -Filename $fieldMappingsFile
	$xmlRoot = $xmlDoc.DocumentElement

	$xpath = "//FieldMappings/Field[@DisplayName='$DisplayName']"
	$node = $xmlRoot.SelectSingleNode($xpath)

	$result = $node.InternalName

	return $result
}

function TestThresholdForSku
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[Parameter(Mandatory=$true)]
		[int]$AcvailableUnits,
		[Parameter(Mandatory=$true)]
		[int]$CurrentUnits,
		[string]$ThresholdFile = "$PSScriptRoot\SkuThresholds.xml"
	)

	$result = 0

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
		$threshold = $node.Threshold

		if ($threshold -ne "0")
		{
			$thresholdValue = [int]$threshold
			$diff = $AcvailableUnits - $CurrentUnits

			if ($diff -le $thresholdValue)
			{
				$result = $diff
			}
		}
	}

	return $result
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
Import-Module PBSPOPS

if ((IsMsolServiceConnected) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a Msol Service. Processing stopped."
	Exit
}

if ((Test-PBSPOConnection) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a SharePoint Online site collection. Processing stopped."
	Exit
}

$transcriptExtension = Get-Date -Format yyyyMMdd-HHmmss
$transcriptFile = "$TranscriptPath\SAP_PrepareFieldsForSku_Transcript_$transcriptExtension.txt"

Start-Transcript -Path $transcriptFile

$currentTimestamp = ([DateTime]::Now).ToString("dd.MM.yyyy")

$accountSkuCollection = Get-MsolAccountSku

Write-Host "Found " + $accountSkuCollection.Count + " items for account sku"

$hashValues = @{}
$hashValues.Add("Title", $currentTimestamp)
$hashValues.Add("CheckDate", [DateTime]::Now)

$skuToNotify = @{}

foreach ($accountSku in $accountSkuCollection)
{
	$skuPartNumber = $accountSku.SkuPartNumber

	Write-Host "Processing sku $skuPartNumber"

	$availableUnits = [int]$accountSku.ActiveUnits
	$currentUnits = [int]$accountSku.ConsumedUnits

	$availableField = "$skuPartNumber Available"
	$availableFieldInternal = GetInternalFieldname -DisplayName $availableField
	$hashValues.Add($availableFieldInternal, $availableUnits)

	$currentField = "$skuPartNumber Current"
	$currentFieldInternal = GetInternalFieldname -DisplayName $currentField
	$hashValues.Add($currentFieldInternal, $currentUnits)

	$thresholdTest = TestThresholdForSku -Sku $skuPartNumber -AcvailableUnits $availableUnits -CurrentUnits $currentUnits 

	if ($thresholdTest -gt 0)
	{
		$skuToNotify.Add($skuPartNumber, $thresholdTest)

		Write-Host "Marked $skuPartNumber for notification mail"
	}
}

$item = Add-PnPListItem -List $LICENSE_TRACKING_LIST -ContentType 0x01 -Values $hashValues

Write-Host "Item added to the license tracking list."

if ($skuToNotify.Count -eq 0)
{
	Write-Host "No need to send any notification mails."
}
else
{
	Write-Host "Start sending notification mails."

	$receipient = GetDefaultReceipient

	. .\SendMail.ps1 -MailType SkuAlert -SkuToNotify $skuToNotify -Receipient $receipient

	<#
	foreach ($sku in $skuToNotify.Keys)
	{
		$currentValue = $skuToNotify[$sku]
		$receipient = GetReceipientForSku -Sku $sku

		Write-Host -NoNewline "Send mail to $receipient "
		Write-Host -ForegroundColor Magenta "=> TODO"
	}
	#>
}

Write-Host -ForegroundColor Green "Done."

Stop-Transcript
