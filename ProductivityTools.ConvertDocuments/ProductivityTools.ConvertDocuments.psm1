
function ExtractPanDoc{
	[CmdletBinding()]
	param()

	$path="$PSScriptRoot\Pandoc.zip"
	Write-Verbose "Trying to extract zip $path"

	Expand-Archive -Path $path -DestinationPath $PSScriptRoot
	
}

function ValidatePandocExecutable{
	[CmdletBinding()]
	param()

	if(Test-Path "$PSScriptRoot\Pandoc\pandoc.exe")
	{
		Write-Verbose "Pandoc.exe exists, we can proceed!"
	}
	else {
		Write-Verbose "Pandoc.exe doesn't exist we need to extract it"
		ExtractPanDoc
	}
}

function ConvertFrom{
	param(
		[Parameter(Mandatory=$true)]
		[string]$Path,

		[Parameter(Mandatory=$true)]
		[string]$TargetFormat,
		
		[string]$OutputDirectory,
		[string]$OutputFileName,
		[string]$SourceFormat,
		
		[switch]$ExtractImages
	)
	
	ValidatePandocExecutable

	Write-Verbose "Path to docx document: $Path"
	Write-Verbose "Chosen target format $TargetFormat"
	
	if($OutputDirectory -eq '')
	{
		$OutputDirectory ='.\';
		Write-Verbose "Output directory empty so changed to $OutputDirectory"
	}
	Write-Verbose "Output directory $OutputDirectory"
	New-Item -ItemType Directory -Force -Path $OutputDirectory
	
	$extension='.txt'
	if ($TargetFormat -eq 'html')
	{
		$extension='html'	
	}
	
	if ($TargetFormat -eq 'Markdown')
	{
		$extension='md'	
	}
	
	if($OutputFileName -eq '')
	{
		$OutputFileName ="Result.$extension";
		Write-Verbose "OutputFileName empty so changed to $OutputFilename"
	}
	
	Write-Verbose "Output OutputFileName: $OutputFileName"
	
	$OutputFullName=Join-Path $OutputDirectory $OutputFileName
	Write-Verbose "Output FullName: $OutputFullName"
	
	$OutputFilesPath=Join-Path $OutputDirectory "Files"

	$scriptPath= $PSScriptRoot
	$pandocPath="$scriptPath\Pandoc\pandoc.exe"

	Write-Verbose "Pandoc path: $pandocPath"
	if($ExtractImages)
	{
		$expression="'$pandocPath' '$Path' -f $SourceFormat -t $TargetFormat --output '$OutputFullName' --extract-media='$OutputFilesPath'"
	}
	else
	{
		[Array]$arguments = $Path, "-f" ,$SourceFormat, "-t", $TargetFormat, "--output", $OutputFullName;
	}
	Write-Verbose  "Executing command: $pandocPath"
	Write-Verbose  "With arguments: $arguments"

	& $pandocPath $arguments
	
}


function ConvertFrom-Markdown{
	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$true)]
		[string]$Path,

		[Parameter(Mandatory=$true)]
		[ValidateSet("Html", "Docx")]
		[string]$TargetFormat,
		
		[string]$OutputDirectory,
		[string]$OutputFileName,
		
		[switch]$ExtractImages
	)
	
	Write-Verbose "Hello from ConvertFrom-Markdown"
	ConvertFrom $Path $TargetFormat $OutputDirectory $OutputFileName "Markdown" -ExtractImages:$ExtractImages
}

function ConvertFrom-Docx {
	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$true)]
		[string]$Path,

		[Parameter(Mandatory=$true)]
		[ValidateSet("Html", "Markdown")]
		[string]$TargetFormat,
		
		[string]$OutputDirectory,
		[string]$OutputFileName,
		
		[switch]$ExtractImages
	)
	

	Write-Verbose "Hello from ConvertFrom-Docx"
	ConvertFrom $Path $TargetFormat $OutputDirectory $OutputFileName "Docx"  -ExtractImages:$ExtractImages
}

Export-ModuleMember ConvertFrom-Docx
Export-ModuleMember ConvertFrom-Markdown