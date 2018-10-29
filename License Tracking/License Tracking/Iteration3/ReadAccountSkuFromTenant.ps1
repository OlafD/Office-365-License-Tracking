param (
	[string]$Listname = "License Tracking",
	[string]$ListnamePerSku = "License Tracking per Sku",
	[Parameter(Mandatory=$true)]
	[string]$TranscriptPath,
	[Parameter(Mandatory=$true)]
	[string]$DefaultReceipient,
	[Parameter(Mandatory=$true)]
	$Credentials
)

#------- Constants -------



#------- Functions -------

<#
 test the connection the the MSOL Service
#>
function IsMsolServiceConnected
{
    $values = Get-MsolAccountSku -ErrorAction SilentlyContinue

	return ($values -ne $null)
}

<#
 test the connection to SharePoint Online
#>
Function TestSPOConnection
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

<#
 get the internal name for a field by the DisplayName of the field
#>
function GetInternalFieldname
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$DisplayName
	)

	$result = ""

	$camlQuery = "<Eq><FieldRef Name='DisplayName' /><Value Type='Text'>$DisplayName</Value></Eq>"

	$item = GetListItems -Listname "Field Mappings" -WhereNode $camlQuery

	if ($item -ne $null)
	{
		$result = $item["InternalName"]
	}

	return $result
}

<#
 test, whether the available units are below the threshold set for an Sku
#>
function TestThresholdForSku
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[Parameter(Mandatory=$true)]
		[int]$AcvailableUnits,
		[Parameter(Mandatory=$true)]
		[int]$CurrentUnits
	)

	$result = 0

	$camlQuery = "<Eq><FieldRef Name='Title' /><Value Type='Text'>$Sku</Value></Eq>"

	$item = GetListItems -Listname "Sku Thresholds" -WhereNode $camlQuery

	if ($item -eq $null)
	{
		Write-Host -ForegroundColor Red "Sku $Sku not found in threshold list."
	}
	else
	{
		$threshold = $item["Threshold"]

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

<#
 Get the configured threshold for the Sku
#>
function GetThresholdForSku
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku
	)

	$result = 0

	$camlQuery = "<Eq><FieldRef Name='Title' /><Value Type='Text'>$Sku</Value></Eq>"

	$item = GetListItems -Listname "Sku Thresholds" -WhereNode $camlQuery

	if ($item -eq $null)
	{
		Write-Host -ForegroundColor Red "Sku $Sku not found in threshold list."
	}
	else
	{
		$result = $item["Threshold"]
	}

	return $result
}

<#
 return the default receipient configured for this tool
#>
function GetDefaultReceipient
{
	return $DefaultReceipient
}

<#
 Get the reciepient for an Sku from the "Sku Thresholds" list
#>
function GetReceipientForSku
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[string]$DefaultReceipient
	)

	$result = ""

	$camlQuery = "<Eq><FieldRef Name='Title' /><Value Type'Text'>$Sku</Value></Eq>"

	$items = GetListItems -Listname "Sku Thresholds" -WhereNode $camlQuery

	if ($items -ne $null)
	{
		$result = $items["Receipient"]

		if ($result -eq "")
		{
			$result = $DefaultReceipient
		}
	}

	return $result
}

<#
 run a query on a list an return the result set
#>
function GetListItems
{
	param (
		[string]$Listname,
		[string]$WhereNode
	)

	$camlQuery = "<View><Query><Where>$WhereNode</Where></Query></View>"

	$result = Get-PnPListItem -List $Listname -Query $camlQuery

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
$transcriptFile = "$TranscriptPath\SAP_ReadAccountSkuFromTenant_Transcript_$transcriptExtension.txt"

Start-Transcript -Path $transcriptFile

$currentTimestamp = ([DateTime]::Now).ToString("dd.MM.yyyy")

$accountSkuCollection = Get-MsolAccountSku

Write-Host "Found" $accountSkuCollection.Count "items for account sku"

$checkDate = [DateTime]::Now
$hashValues = @{}
$hashValues.Add("Title", $currentTimestamp)
$hashValues.Add("CheckDate", $checkDate)

$skuToNotify = @{}

foreach ($accountSku in $accountSkuCollection)
{
	$skuPartNumber = $accountSku.SkuPartNumber

	Write-Host "Processing sku $skuPartNumber"

	$availableUnits = [int]$accountSku.ActiveUnits
	$currentUnits = [int]$accountSku.ConsumedUnits

#region List 'License Tracking'
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
#endregion

#region List 'License Tracking per Sku'
	$threshold = GetThresholdForSku -Sku $skuPartNumber
	$hash = @{}
	$hash.Add("Title", $skuPartNumber)
	$hash.Add("CheckDate", $checkDate)
	$hash.Add("Available", $availableUnits)
	$hash.Add("Current", $currentUnits)
	$hash.Add("Threshold", $threshold)

	$item = Add-PnPListItem -List $ListnamePerSku -ContentType 0x01 -Values $hash

	Write-Host "Item for Sku $skuPartNumber added to list '$ListnamePerSku'."
#endregion
}

$item = Add-PnPListItem -List $Listname -ContentType 0x01 -Values $hashValues

Write-Host "Item added to the license tracking list."

if ($skuToNotify.Count -eq 0)
{
	Write-Host "No need to send any notification mails."
}
else
{
	Write-Host "Start sending notification mails."

	$receipient = GetDefaultReceipient

	. .\SendMail.ps1 -MailType SkuAlert -SkuToNotify $skuToNotify -Credentials $Credentials
}

Write-Host -ForegroundColor Green "Done."

Stop-Transcript

Write-Output $item.Id
