
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

Function Rename-Field
{
<#
	.SYNOPSIS
	Rename the title of the field in the default language of the site

	.DESCRIPTION
	Set the title property of the field to a new value. This will change the display name of the
	field.

	.PARAMETER Identity
	The title, the internal name or the id of the field, for which the title should be renamed.

	.PARAMETER List
	The name of the list, in which the field could be found. If this parameter is empty, the field
	is taken from the site columns

	.PARAMETER NewValue
	The new title for the field

	.NOTES
	no notes available

	.LINK
	no link available

	.EXAMPLE
	Rename-Field -Identity "Title" -NewValue "Company"

	.EXAMPLE
	Rename-Field -Identity "Title" -NewValue "Company" -List "Companies"

#>

	param(
		[Parameter(Mandatory=$true)]
		[string]$Identity,
		[Parameter(Mandatory=$true)]
		[string]$NewValue,
		[Parameter(Mandatory=$false)]
		[string]$List,
		[switch]$PushUpdate
	)

	if ((TestSPOConnection) -eq $false)
	{
		Write-Error "No connection to SharePoint"

		break
	}

	if ($List -eq "")  # get field from site columns
	{
		$field = Get-PnPField -Identity $Identity
	}
	else
	{
		$field = Get-PnPField -Identity $Identity -List $List
	}

	if ($field -ne $null)
	{
		$field.Title = $NewValue

		if ($PushUpdate.ToBool() -eq $true)
		{
			Write-Host "Propagate changes..."

			$field.UpdateAndPushChanges($true)
		}
		else
		{
			$field.Update()
		}

		$ctx = Get-PnPContext
		$ctx.ExecuteQuery()
	}
}

Function Add-FieldToList()
{
<#
	.SYNOPSIS
	Add a field to the columns in a list.

	.DESCRIPTION
	This function uses the cmdlet Add-PnPFieldFromXml from the Office 365 PnP PowerShell extension
	to add a new field to a list. The parameters for the field are taken from an xml-file, the 
	filename is passed as the parameter.

	.PARAMETER List
	The name of the list, where the new field should be added

	.PARAMETER Path
	The path to the xml-file with the parameters for the new field

	.NOTES
	no notes available

	.LINK
	no link available

	.EXAMPLE
	$field = Add-FieldToList -List "Documents" -Path "C:\Columns\Status.xml"

#>

	param(
		[Parameter(Mandatory=$true)]
		[string]$List,
		[Parameter(Mandatory=$true)]
		[string]$Path
	)

	if ((TestSPOConnection) -eq $false)
	{
		Write-Error "No connection to SharePoint"

		break
	}

	$content = Get-Content $Path

	$contentAsString = [string]$content

	$field = Add-PnPFieldFromXml -List $List -FieldXml $contentAsString

	return $field
}

Function Set-ShowInForm
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Listname,
		[Parameter(Mandatory=$true)]
		[string]$Identity,
		[Parameter(Mandatory=$true)]
		[ValidateSet("New", "Edit", "Display")]
		[string]$FormType,
		[Parameter(Mandatory=$true)]
		[bool]$Value
	)

	$field = Get-PnPField -List $Listname -Identity $Identity -ErrorAction SilentlyContinue

	if ($field -eq $null)
	{
		Write-Host -ForegroundColor Red "Cannot find field '$Identity' in list '$Listname'."
	}
	else
	{
		switch ($FormType)
		{
			"New" 
			{
				$field.SetShowInNewForm($Value)
				$field.Update()
				Invoke-PnPQuery
			}
			"Edit" { 
				$field.SetShowInEditForm($Value)
				$field.Update()
				Invoke-PnPQuery
			}
			"Display" {
				$field.SetShowInDisplayForm($Value)
				$field.Update()
				Invoke-PnPQuery
			}
		}
	}
}