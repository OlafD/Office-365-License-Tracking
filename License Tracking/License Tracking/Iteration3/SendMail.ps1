param (
       [ValidateSet("SkuAlert","NewSku")]
       [string]$MailType,
       [Parameter(Mandatory=$true)]
       [Hashtable]$SkuToNotify,
       [Parameter(Mandatory=$true)]
       $Credentials
)
 
$NEW_SKU_SUBJECT = "New Sku in license alerting"
$NEW_SKU_BODY = @"
<div>Hello,</div>
<br />
<div>this automatic E-mail is send to you as an alert notification for a new Sku in the SkuThresholds.xml for the license alerting. The following Sku(s) were added:</div>
<br />
<ul>
[*license_placeholder*]
</ul>
<br />
<div>Please add the friendly name for the new Sku to the xml file to make it more readable for the alert receipients.</div>
<br />
<br />
<div>Thank you and Kind regards</div>
<br />
<div>The Office 365 License Management Service.</div>
"@
 
$SKU_ALERT_SUBJECT = "Office 365 License Alert"
$SKU_ALERT_BODY = @"
<div>Hello,</div>
<br />
<div>this automatic E-mail is send to you as an alert notification of the threshold set for the following Office 365 licenses.</div>
<br />
<ul>
[*license_placeholder*]
</ul>
<br />
<div>Please take appropriate actions for future operational assurance of the service.</div>
<br />
<br />
<div>Thank you and Kind regards</div>
<br />
<div>The Office 365 License Management Service.</div>
"@
 
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
 
function GetValueFromXml
{
       param (
             [Parameter(Mandatory=$true)]
             [string]$NodeName,
             [string]$XmlFilename = "$PSScriptRoot\LicenseTrackingParam.xml"
       )
 
       $result = ""
 
       $xmlDoc = New-Object System.Xml.XmlDocument
       $xmlDoc.Load($XmlFilename)
       $rootNode = $xmlDoc.DocumentElement
 
       $node = $rootNode.SelectSingleNode("//LicenseTracking/$NodeName")
 
       if ($node -ne $null)
       {
             $result = $node.InnerText
       }
 
       return $result
}
 
function GetNodeFromXml
{
	param (
			[Parameter(Mandatory=$true)]
			[string]$NodeName,
			[string]$XmlFilename = "$PSScriptRoot\LicenseTrackingParam.xml"
	)

	$result = $null
	  
	$xmlDoc = New-Object System.Xml.XmlDocument
	$xmlDoc.Load($XmlFilename)
	$rootNode = $xmlDoc.DocumentElement
 
	$node = $rootNode.SelectSingleNode("//LicenseTracking/$NodeName")
 
	if ($node -ne $null)
	{
			$result = $node
	}
 
	return $result
}
 
<#
prepare the placeholder for the Skus for the notification mail
#>
function CreateSkuAlertPlaceholder
{
       param (
             [Parameter(Mandatory=$true)]
             [Hashtable]$SkuToNotify
       )
 
       $result = ""
 
       foreach ($sku in $SkuToNotify.Keys)
       {
             $skuFriendlyName = GetFriendlyNameForSku -Sku $sku
             $skuThreshold = GetThresholdForSku -Sku $sku
             $skuValue = $SkuToNotify[$sku]
 
             $line = "$skuFriendlyName ($sku) - The threshold, set at $skuThreshold, has been exceeded and <span style='font-weight: bold;'>the current amount of available licenses is only $skuValue.</span>"
             $result += "<li>$line</li>"
       }
 
       return $result
}
 
 
<#
prepare the placeholder for new Skus for the notification mail
#>
function CreateNewSkuPlaceholder
{
       param (
             [Parameter(Mandatory=$true)]
             [Hashtable]$SkuToNotify
       )
 
       $result = ""
 
       foreach ($sku in $SkuToNotify.Keys)
       {
             $line = "$sku"
             $result += "<li>$line</li>"
       }
 
       return $result
}
 
<#
for an Sku get the configured threshold
#>
function GetThresholdForSku
{
       param (
             [Parameter(Mandatory=$true)]
             [string]$Sku,
             [string]$ThresholdFile = "$PSScriptRoot\SkuThresholds.xml"
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
             $result = [int]$item["Threshold"]
       }
 
       return $result
}
 
<#
for an Sku get the configured friendly name
#>
function GetFriendlyNameForSku
{
       param (
             [Parameter(Mandatory=$true)]
             [string]$Sku,
             [string]$ThresholdFile = "$PSScriptRoot\SkuThresholds.xml"
       )
 
       $result = $Sku
 
       $camlQuery = "<Eq><FieldRef Name='Title' /><Value Type='Text'>$Sku</Value></Eq>"
 
       $item = GetListItems -Listname "Sku Thresholds" -WhereNode $camlQuery
 
       if ($item -eq $null)
       {
             Write-Host -ForegroundColor Red "Sku $Sku not found in threshold list."
       }
       else
       {
             $result = $item["FriendlyName"]
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

<#
  build the array for the receipients
#>
function GetReceipients
{
	$result = $null

	$receipients = GetNodeFromXml -NodeName "Receipients"

	if ($receipients -eq $null)
	{
		Write-Host -ForegroundColor Red "No receipients found in the LicenseTrackingParam.xml"
	}
	else
	{
		$nodeCount = $receipients.ChildNodes.Count
		$result = New-Object -TypeName 'string[]' -ArgumentList $nodeCount

		for ($i=0; $i -lt $nodeCount; $i++)
		{
			$node = $receipients.ChildNodes[$i]

			$result[$i] = $node.InnerText
		}
	}

	return $result
} 

#------- Main -------
 
if ((TestSPOConnection) -eq $false)
{
       Write-Host -ForegroundColor Red "No connection to a SharePoint Online site collection. Processing stopped."
       Exit
}
 
switch ($MailType)
{
       "SkuAlert"
       {
             $placeholderValue = CreateSkuAlertPlaceholder -SkuToNotify $SkuToNotify
             $body = $SKU_ALERT_BODY.Replace("[*license_placeholder*]", $placeholderValue)
             $subject = $SKU_ALERT_SUBJECT
             break;
       }
       "NewSku"
       {
             $placeholderValue = CreateNewSkuPlaceholder -SkuToNotify $SkuToNotify
             $body = $NEW_SKU_BODY.Replace("[*license_placeholder*]", $placeholderValue)
             $subject = $NEW_SKU_SUBJECT
             break;
       }
}
 
Write-Host "Send mail to $Receipient"
 
$smtpServer = GetValueFromXml -NodeName "SmtpServer"
$smtpPort = GetValueFromXml -NodeName "SmtpPort"
$from = GetValueFromXml "MailFrom"
$receipients = GetReceipients

$receipients

if ($receipients.Count -gt 0)
{
	Send-MailMessage -From $from -To $receipients -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -Port $smtpPort -Credential $Credentials -UseSsl
} 
else
{
	Write-Host -ForegroundColor Red "No receipients defined."
}
