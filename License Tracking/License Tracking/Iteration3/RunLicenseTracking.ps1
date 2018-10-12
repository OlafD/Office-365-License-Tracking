
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

$url = GetValueFromXml -NodeName "Url"
$listname = GetValueFromXml -NodeName "Listname"
$listnamePerSku = GetValueFromXml -NodeName "ListnamePerSku"
$transcriptPath = GetValueFromXml -NodeName "TranscriptPath"
$defaultReceipient = GetValueFromXml -NodeName "DefaultReceipient"

# for unattended execution, add a mechanism to create the credential object
$cred = Get-Credential 

# connect to the necessary services

Connect-PnPOnline -Url $url -Credentials $cred

Connect-MsolService -Credential $cred

# run the scripts for schema changes and the license tracking itself

# the script will use the default values for the list to store the values, so the parameter Listname is not needed
. .\PrepareFieldsForSku.ps1 -TranscriptPath $transcriptPath -DefaultReceipient $defaultReceipient -Credentials $cred

# the script will use the default values for the lists to store the values, so the parameters Listname and ListnamePerSku are not needed
$itemId = . .\ReadAccountSkuFromTenant.ps1 -TranscriptPath $transcriptPath -DefaultReceipient $defaultReceipient -Credentials $cred

Write-Host -ForegroundColor Green "Done."
