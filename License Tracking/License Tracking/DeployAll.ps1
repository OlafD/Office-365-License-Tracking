param (
	[Parameter(Mandatory=$true)]
	[string]
	$Url,
	$Credentials,
	[switch]
	$UseWebLogin
)

Write-Host -ForegroundColor Magenta "Deploy the License Tracking structure to $Url"

#region Connect to SharePoint site

Write-Host -ForegroundColor Yellow "Connect to $url"

if (($UseWebLogin.ToBool() -eq $false) -and ($Credentials -eq $null))
{
	$Credentials = Get-Credential
}

if ($UseWebLogin.ToBool() -eq $false)
{
	Connect-PnPOnline -Url $Url -Credentials $Credentials
}
else
{
	Connect-PnPOnline -Url $Url -UseWebLogin
}

#endregion

#region Iterations section

Write-Host "Iteration1"

cd .\Iteration1

. .\CreateList.ps1

Write-Host "Iteration2"

cd ..\Iteration1

. .\CreateLists.ps1

#endregion

Write-Host -ForegroundColor Green "Done."
