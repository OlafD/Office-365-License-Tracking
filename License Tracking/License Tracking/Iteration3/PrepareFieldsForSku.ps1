param (
	[string]$Listname = "License Tracking",
	[Parameter(Mandatory=$true)]
	[string]$TranscriptPath,
	[Parameter(Mandatory=$true)]
	[string]$DefaultReceipient,
	[Parameter(Mandatory=$true)]
	$Credentials
)

#------- Constants -------

$LICENSE_TRACKING_LIST = $Listname


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

<#
 load an xml file and return the XmlDocument object
#>
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
 from the list "Field Mappings" find the highest value for the field "FieldNumber"
#>
function FindLastFieldNumber
{
	$result = 0

	$camlQuery = '<View><Query><OrderBy><FieldRef Name="FieldNumber" Ascending="FALSE" /></OrderBy></Query></View>'

	$items = GetListItems -Listname "Field Mappings" -ViewNode $camlQuery

	if ($items.Count -gt 0)
	{
		$item = $items[0]
		$result = [int]$item["FieldNumber"]
	}

	return $result
}

<#
 Check, whether the Sku is already in the list "Field Mappings"
#>
function IsSkuInFieldMappings
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku
	)

	$result = $false

	$camlQuery = "<Eq><FieldRef Name='Title' /><Value Type='Text'>$Sku</Value></Eq>"

	$items = GetListItems -Listname "Field Mappings" -WhereNode $camlQuery

	if ($items -ne $null)
	{
		$result = $true
	}

	return $result
}

<#
 Add a new item for an Sku to the list "Field Mappings"
#>
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
		[int]$FieldNumber
	)

	$hash = @{}
	$hash.Add("Title", $Sku)
	$hash.Add("DisplayName", $DisplayName)
	$hash.Add("InternalName", $InternalName)
	$hash.Add("FieldId", $Id)
	$hash.Add("FieldNumber", $FieldNumber)

	$i = Add-PnPListItem -List "Field Mappings" -ContentType 0x01 -Values $hash
}

<#
 Add a new item for a Sku to the list "Sku Thresholds"
#>
function AddSkuToThresholds
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Sku,
		[Parameter(Mandatory=$true)]
		[string]$Receipient
	)

	$hash = @{}
	$hash.Add("Title", $Sku)
	$hash.Add("FriendlyName", $Sku)
	$hash.Add("Threshold", 0)
	$hash.Add("Receipient", $Receipient)

	$i = Add-PnPListItem -List "Sku Thresholds" -ContentType 0x01 -Values $hash

}

<#
 Add a new field for an Sku to the "License Tracking" list
#>
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
			AddFieldToDefaultView -InternalName $InternalName -Listname $LICENSE_TRACKING_LIST
		}
	}
}

<#
 return the default receipient configured for this tool
#>
function GetDefaultReceipient
{
	return $DefaultReceipient
}

<#
 Add a field to the default view of a list
#>
function AddFieldToDefaultView
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Listname = $LICENSE_TRACKING_LIST,
		[Parameter(Mandatory=$true)]
		[string]$InternalName
	)

	$view = Get-PnPView -List $Listname -Includes ViewFields
	$view.ViewFields.Add($InternalName)
	$view.Update()
	Invoke-PnPQuery
}

<#
 Create a new internal name for a field
#>
function CreateInternalName
{
	param (
		[Parameter(Mandatory=$true)]
		[int]$Number		
	)

	return "Field" + $Number.ToString("000")
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
		[string]$WhereNode,
        [string]$ViewNode
	)

    if ($ViewNode -ne "")
    {
        $camlQuery = $ViewNode
    }
    else
    {
    	$camlQuery = "<View><Query><Where>$WhereNode</Where></Query></View>"
    }

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

		Write-Host "- added $displayName to FieldMapping list"

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

		Write-Host "- added $displayName to FieldMapping list"

		# add sku to Threshold file
		AddSkuToThresholds -Sku $skuPartNumber -Receipient $DefaultReceipient

		$newSkuNotification.Add($skuPartNumber, $skuPartNumber)

		Write-Host "- added $skuPartNumber to thresholds list"
	}
    else
    {
		Write-Host -ForegroundColor Green "- found in field mappings, nothing to do"
    }
}

if ($newSkuNotification.Count -gt 0)
{
	$receipient = GetDefaultReceipient

	. .\SendMail.ps1 -MailType NewSku -SkuToNotify $newSkuNotification -Credentials $Credentials
}

Write-Host -ForegroundColor Green "Done."

Stop-Transcript
