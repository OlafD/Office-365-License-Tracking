
. ..\PowerShellFunc\PowerShellFunc.ps1

if ((TestSPOConnection) -eq $false)
{
	Write-Host -ForegroundColor Red "No connection to a SharePoint Online site collection. Processing stopped."
	
	Exit
}

Write-Host -ForegroundColor Yellow "Create list 'License Tracking per Sku'"

$listname = "License Tracking per Sku"

$l = New-PnPList -Title $listname -Template GenericList -Url "licensetrackingsku"

$list = Get-PnPList -Identity $listname -Includes "ListExperienceOptions"

$list.ListExperienceOptions = "ClassicExperience"
$list.Update()
Invoke-PnPQuery

$f = Add-FieldToList -Path .\Columns\CheckDate.xml -List $listname
$f = Add-FieldToList -Path .\Columns\Available.xml -List $listname
$f = Add-FieldToList -Path .\Columns\Current.xml -List $listname
$f = Add-FieldToList -Path .\Columns\Threshold.xml -List $listname

$view = Get-PnPView -List $listname -Includes ViewFields
$view.ViewFields.Add("CheckDate")
$view.ViewFields.Add("Available")
$view.ViewFields.Add("Current")
$view.ViewFields.Add("Threshold")
$view.Update()
Invoke-PnPQuery

Set-ShowInForm -List $listname -Identity "Title" -FormType "New" -Value $false
Set-ShowInForm -List $listname -Identity "Title" -FormType "Edit" -Value $false

Set-ShowInForm -List $listname -Identity "CheckDate" -FormType "New" -Value $false
Set-ShowInForm -List $listname -Identity "CheckDate" -FormType "Edit" -Value $false

Set-ShowInForm -List $listname -Identity "Available" -FormType "New" -Value $false
Set-ShowInForm -List $listname -Identity "Available" -FormType "Edit" -Value $false

Set-ShowInForm -List $listname -Identity "Current" -FormType "New" -Value $false
Set-ShowInForm -List $listname -Identity "Current" -FormType "Edit" -Value $false

Set-ShowInForm -List $listname -Identity "Threshold" -FormType "New" -Value $false
Set-ShowInForm -List $listname -Identity "Threshold" -FormType "Edit" -Value $false


Write-Host -ForegroundColor Green "Done."
