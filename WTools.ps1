<#
.SYNOPSIS
This is a library of helpful functions and variables.

.DESCRIPTION
This library was intended to provide a common set of functions and variables.  This script should be dot-sourced in scripts so that all functions/variables can be updated from one central source.

.EXAMPLE
. \\myServer\scriptShare\wtools.ps1

.NOTES
v0.2        HeyDude378      alpha build 2
v0.1		HeyDude378		alpha build 1

.COMPONENT
Requires .NET installed on host OS.
Requires ActiveDirectory PowerShell module.
#>

#Region EnvironmentVariables
$infastructureMaster = "$((get-addomain (get-adforest).name).infrastructuremaster)" + ":3268"
#EndRegion EnvironmentVariables

#Region Functions
function New-RandomPassword {
	<#
	.SYNOPSIS
	Creates a password of the specified length.

	.DESCRIPTION
	Valid characters are selected from all allowed symbols, integers, capital letters, and lowercase letters, excluding: 0Oo1lIi

	.PARAMETER LENGTH
	A valid password of the specified number of characters will be created.  Range is from 1 to 127.

	.EXAMPLE
	New-RandomPassword -length 8

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Returns a string value.

	.NOTES
	v1.0		HeyDude378		initial stable build

	#>

	param (
		[Parameter(Mandatory)]
		[ValidateRange(1, 127)][int] $length
	)

    if($length -eq 0){
        $length=8
        Write-Host -foregroundcolor Yellow "Using default length of 8!"
    }

    $capitals = [char[]] "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$lowercases = [char[]] "abcdefghijklmnopqrstuvwxyz" 
	$numbers = [char[]] "0123456789"
	$specials = [char[]] "~!@#$%^&*_-+=``|\(){}[]:;`"`'<>,.?/" #all special characters that are valid for passwords
	$exclusions = [char[]] "0Oo1lIi" #characters that look too much like each other 
	
	$validCharacters = New-Object -TypeName "System.Collections.ArrayList"
	$validCharacters.AddRange($capitals) | Out-Null
	$validCharacters.AddRange($lowercases) | Out-Null
	$validCharacters.AddRange($numbers) | Out-Null
	$validCharacters.AddRange($specials) | Out-Null
	$exclusions | ForEach-Object {$validCharacters.Remove($_) | Out-Null}
	
	$i=0
	do {$pwdString += ($validCharacters | get-random -count 1);$i++}
		while ($i -lt $length)
	
	return $pwdString
}

function Select-File {
	<#
	.SYNOPSIS
	File-explorer GUI based file picker.

	.DESCRIPTION
	Opens an OpenFileDialog defaulting to user profile folder.

	.EXAMPLE
	$file = Select-File

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Outputs a string corresponding to the full path and file name of the selected file.

	.NOTES
	v1.0		HeyDude378		initial stable build

	.LINK
	(none)

	.COMPONENT
	Requires .NET libraries.

	.FUNCTIONALITY
	pick, select, file, GUI, explorer
	#>

	Write-Host -foregroundcolor Yellow "`n`nAttempting to import a file.  If you do not see a file dialog, it may be behind another window.  This often happens when executing scripts from Visual Studio Code."
    Add-Type -AssemblyName System.Windows.Forms
    $null = $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
    }
    $FileBrowser.showdialog() | Out-Null
    return $FileBrowser.FileName
}

function Search-ArrayList {
<#
	.SYNOPSIS
	Searches an arraylist for values from specified fields.

	.DESCRIPTION
	Searches each field of the array for the specified text.  Returns all matching elements of the array, or if the SingleResult switch is used, disambiguates to one result.

	.PARAMETER SearchArray
    Expects an array or arraylist.  Use with a predefined array or arraylist, or use Import-CSV to input at runtime, e.g. Search-ArrayList -SearchArray (Import-CSV -Path $env:userprofile\desktop\myArray.csv) -SearchField "Color" -SearchTerm "Green"

    .PARAMETER SingleResult
    Use this switch to specify that you only want one result.  If more than one is found, you will choose which one you want.

    .PARAMETER SearchTerm
    The text you wish to search for.

    .PARAMETER SearchField
    The field (i.e. property) you wish to search in.

	.EXAMPLE
	Search-ArrayList -searchArray $myProjectSpreadsheet -SearchField "Project Status" -SearchTerm "In Progress"

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Outputs matched elements from the input arraylist.

	.NOTES
	v1.0		HeyDude378		initial stable build

	.LINK
	(none)

	.COMPONENT
	Requires .NET libraries.

	.FUNCTIONALITY
	CSV, spreadsheet, search, find, arraylist
	#>

	param(
		[System.Collections.ArrayList]$SearchArray,
		[switch]$SingleResult,
		$SearchTerm,
		$SearchField
	)

	if([string]::IsNullOrEmpty($SearchTerm)){[string]$SearchTerm = Read-Host "Please enter text to search for"}
	$result = New-Object System.Collections.ArrayList

	($SearchArray | Where-Object $SearchField -like "*$($SearchTerm)*") | foreach-object{$result.Add($_) | Out-Null}

	if($result.count -eq 0){
		Write-Host -ForegroundColor Red "No results contain the value: $($SearchTerm)"
	}
	else{
		if(($result.count -gt 1) -and ($SingleResult)){
			$i=0
			$result | ForEach-Object {
				$i++
				Write-Host -ForegroundColor Yellow "(Result $i)`: " -NoNewline
				Write-Host "$_"
			}
			[int]$oneOfMany = Read-Host "`nYou may choose only one result.  Enter number to select result"
			$result = $result[$oneofMany-1]
		}		
		return ,$result
	}
}

function Import-ValidCSV {
	<#
	.SYNOPSIS
	Imports a CSV and optionally checks if required fields are contained in the CSV.

	.DESCRIPTION
	Imports a CSV file using Select-File function and checks that each required field is specified in the file.  If not, errors and allows retry.

	.PARAMETER requiredFields
	The CSV file will be checked to ensure that it contains these required column headings.

	.PARAMETER outputType
	Specify array or arraylist as the output type.

	.EXAMPLE
	$MyArrayList = Import-ValidCSV -outputType ArrayList

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Outputs an array or arraylist object as specified in outputType parameter.

	.NOTES
	v1.0		HeyDude378		initial stable build

	.LINK
	(none)

	.COMPONENT
	Requires .NET libraries.

	.FUNCTIONALITY
	data validation, CSV, import, export, array, arraylist
	#>

    param ($requiredFields,
	[Parameter(Mandatory)] [ValidateSet("array","ArrayList")] $outputType)

    if($null -ne $requiredFields){
		Write-Host "The following required fields have been specified:`n" $requiredfields -separator "`n" "`n"
    	Pause
	}

    $csvpath = Select-File
	
	if($outputType -eq "array"){
		$csvObject = import-CSV -path $csvpath
	}

	elseif($outputType -eq "ArrayList"){
		$csvObject = New-Object System.Collections.ArrayList
		$csvObject.AddRange((import-csv -path $csvpath))
	}
	
	if($null -ne $requiredFields){
		$requiredfields | foreach-object {
			if(($csvObject | get-member).name -contains $_){Write-Host -foregroundcolor green "Found $_ field"}
			else{
				Write-Host -foregroundcolor red "Didn't find $_ field.  CSV does not contain all required fields."
				if((Read-Host "Try again?  Y/N") -eq "Y"){Import-ValidCSV}
			}
		}
	}

    return ,$csvObject
}

function Send-Email {
	<#
	.SYNOPSIS
	Sends a plaintext or HTML email.

	.DESCRIPTION
	This function uses .NET functionality to construct and send an email message.  If no mail server is specified, defaults to value of $mailServer.

	.PARAMETER requiredFields
	From is required.  You must also specify at least a To, CC, or BCC value.	

	.EXAMPLE
	Send-Email -From "john@contoso.com" -To "jane@contoso.com" -subject "Favorite Color" -body "What is your favorite color?"

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Outputs a success or failure message.

	.NOTES
	v1.0		HeyDude378		initial stable build

	.LINK
	(none)

	.COMPONENT
	Requires .NET libraries.

	.FUNCTIONALITY
	email, message, exchange, send, html
	#>
	
	param(
		[Parameter(Mandatory=$true)]$from,
		$to,
		$cc,
		$bcc,
		$subject,
		$body,
		[switch]$isHTML,
		$sendingServer = "$mailServer"
	)

	if(($null -eq $to) -and ($null -eq $cc) -and ($null -eq $bcc)){Write-Host -ForegroundColor Red "You must specify at least one To, CC, or BCC.  Please try again."}

	else{
		$SMTPMessage = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
		if($isHTML){
			$body = ($body | ConvertTo-Html)
			$SMTPMessage.IsBodyHtml = $true
		}
		if($null -ne $cc){$SMTPMessage.Cc.Add($cc)}
		if($null -ne $bcc){$SMTPMessage.Bcc.Add($bcc)}
		$SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer
		$SMTPClient.Send($SMTPMessage)
	}
}

function Get-ValidADUserObject {
	<#
	.SYNOPSIS
	Gets a valid AD user account object from the current forest.

	.DESCRIPTION
	Checks entire AD forest for user, then disambiguates if necessary, then confirms before sending output.

	.EXAMPLE
	Get-ValidADUserObject | Set-ADUser -City "London"

	.INPUTS
	This function does not accept input from the pipeline.

	.OUTPUTS
	Outputs nothing or an AD user object.

	.NOTES
	v1.0		/u/HeyDude378		initial stable build

	.LINK
	(none)

	.COMPONENT
	Requires ActiveDirectory PowerShell module.

	.FUNCTIONALITY
	find, get, search, AD, Active Directory, user, account
	#>

	param($username)
    
    Clear-Variable userObject -ErrorAction SilentlyContinue
    if($null -eq $username){$username = Read-Host "Enter username"}
    $userObject = Get-ADUser -Filter {samaccountname -eq $username} -Server "$($infrastructureMaster):3268"
    if($null -ne $userObject){
        if($userObject.count -gt 1){
            Write-Warning "Found multiple results:`n`n$($userObject.distinguishedname)"
            $specificDomain = Read-Host "Please type the domain name of the desired user"
            $userObject = Get-ADUser -Identity $username -Server $specificDomain
        }
    }
    Write-Host "Found $($userObject.distinguishedname)."
    $selection = Read-Host "Enter C to Continue, R to Retry, or Q to Quit"
    switch ($selection) {
        "C" {return $userObject}
        "R" {Get-ValidADUserObject}
        default {
            Write-Host "Exiting."
            Clear-Variable userObject -ErrorAction SilentlyContinue
        }
    }
}
#EndRegion Functions

#Region Interface
Write-Host -ForegroundColor Yellow "Successfully loaded WTools!
This script is currently in alpha phase development.  It is being actively developed by HeyDude378.  Type get-help <function name> for more details.`n"

Write-Host "Available Functions:
Get-ValidADUserObject ((v1.0)
Import-ValidCSV (v1.0)
New-RandomPassword (v1.0)
Search-ArrayList (v1.0)
Select-File (v1.0)
Send-Email (v1.0)

Available Environment Variables:
`$im: $im
"
#EndRegion Interface