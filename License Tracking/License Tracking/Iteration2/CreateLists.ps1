
. ..\PowerShellFunc\PowerShellFunc.ps1

Write-Host -ForegroundColor Yellow "Field Mappings"

$listname = "Field Mappings"

$l = New-PnPList -Title $listname -Url "lists/fieldmappings" -Template GenericList -OnQuickLaunch:$false

$list = Get-PnPList -Identity $listname -Includes "ListExperienceOptions"

$list.ListExperienceOptions = "ClassicExperience"
$list.Update()
Invoke-PnPQuery

$f = Rename-Field -List $listname -Identity "Title" -NewValue "SkuPartNumber"

$f = Add-PBFieldToList -Path .\Columns\DisplayName.xml -List $listname
$f = Add-PBFieldToList -Path .\Columns\InternalName.xml -List $listname
$f = Add-PBFieldToList -Path .\Columns\FieldId.xml -List $listname
$f = Add-PBFieldToList -Path .\Columns\FieldNumber.xml -List $listname

$view = Get-PnPView -List $listname -Includes ViewFields
$view.ViewFields.Add("DisplayName")
$view.ViewFields.Add("InternalName")
$view.ViewFields.Add("FieldId")
$view.ViewFields.Add("FieldNumber")
$view.Update()
Invoke-PnPQuery

$titleField = Get-PnPField -List $listname -Identity "Title"
$titleField.SetShowInNewForm($false)
$titleField.SetShowInEditForm($false)
$titleField.Update()
Invoke-PnPQuery


Write-Host -ForegroundColor Yellow "Sku Thresholds"

$listname = "Sku Thresholds"

$l = New-PnPList -Title $listname -Url "lists/skuthresholds" -Template GenericList -OnQuickLaunch:$false

$list = Get-PnPList -Identity $listname -Includes "ListExperienceOptions"

$list.ListExperienceOptions = "ClassicExperience"
$list.Update()
Invoke-PnPQuery

$f = Rename-Field -List $listname -Identity "Title" -NewValue "SkuPartNumber"

$f = Add-PBFieldToList -Path .\Columns\FriendlyName.xml -List $listname
$f = Add-PBFieldToList -Path .\Columns\Threshold.xml -List $listname
$f = Add-PBFieldToList -Path .\Columns\Receipient.xml -List $listname

$view = Get-PnPView -List $listname -Includes ViewFields
$view.ViewFields.Add("FriendlyName")
$view.ViewFields.Add("Threshold")
$view.ViewFields.Add("Receipient")
$view.Update()
Invoke-PnPQuery

$titleField = Get-PnPField -List $listname -Identity "Title"
$titleField.SetShowInNewForm($false)
$titleField.SetShowInEditForm($false)
$titleField.Update()
Invoke-PnPQuery


Write-Host "Done."
