
Import-Module PBSPOPS

if ((Test-PBSPOConnection) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a SharePoint Online site collection. Processing stopped."
	Exit
}

Write-Host -ForegroundColor Yellow "Create list 'License Tracking'"

$listname = "License Tracking"

$l = New-PnPList -Title $listname -Template GenericList -Url "licensetracking"

$list = Get-PnPList -Identity $listname -Includes "ListExperienceOptions"

$list.ListExperienceOptions = "ClassicExperience"
$list.Update()
Invoke-PnPQuery

$f = Add-PBFieldToList -Path .\Columns\CheckDate.xml -List $listname

$view = Get-PnPView -List $listname -Includes ViewFields
$view.ViewFields.Add("CheckDate")
$view.Update()
Invoke-PnPQuery

Write-Host -ForegroundColor Green "Done."
