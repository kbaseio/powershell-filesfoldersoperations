#requires -version 4
<#
.SYNOPSIS
  Get random file from a folder or a file list (one file path per line)

.DESCRIPTION
  Read the content of a folder or a file and select one entry (pseudo)randomely
  You have the options to the copy the file locally (Profile Temp Folder) or / and to open the file once selected

.PARAMETER Path
  Folder Path or File containing the list of file path

.PARAMETER Copy
  Copy the selected file on your computer (Temp Folder)
  To be used if the file os too large or if the software that reads it have troubles with network paths.

.PARAMETER Run
  Open the file 
  If not selected the script will return the file path.

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.0
  Author:         FingersOnFire
  Creation Date:  2021-09-21
  Purpose/Change: Initial script development

.EXAMPLE
  Get-RandomFile -Path C:\Temp
  
  Select a file in C:\Temp and return the path as the output of the script

.EXAMPLE
  Get-RandomFile -Path C:\Temp\List.txt
  
  Select a file from the paths in a TXT file and return the path as the output of the script

.EXAMPLE
  Get-RandomFile -Path C:\Temp -Copy
  
  Select a file in C:\Temp and copy it to the local TEMP folder and return the path as the output of the script


.EXAMPLE
  Get-RandomFile -Path C:\Temp -Copy
  
  Select a file in C:\Temp and copy it to the local TEMP folder and open it in the default associated program


#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

param(
  [Parameter(Mandatory=$true)]
  [string]$Path,

  [Parameter(Mandatory=$false)]
  [Switch]$Copy,

  [Parameter(Mandatory=$false)]
  [Switch]$Run
)



#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$SelectedFile = "" #To be used later
$TempLocation = "$env:TEMP"

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------



if(Test-Path -Path $Path -PathType Container){
    Write-Host "Path is a directory. Getting files within..."
    $SelectedFile = Get-ChildItem $path -File | Get-Random -Count 1
}
elseif(Test-Path -Path $Path -PathType File) {
    Write-Host "Path is a file. Getting content..."
    $SelectedFile = Get-Content $path | Get-Random -Count 1
}
else{
    Write-Host "The given path can not be found. Please call the Ghost Busters..."
    exit
}


#Copy it in a local folder if asked so
if($Copy){
    $TempItem = Copy-Item -Path $SelectedFile -Destination $TempLocation -PassThru

    if(Test-Path $TempItem){
        $SelectedFile = $TempItem
    }
    else{
        Write-Host "Error during local copy... Exiting."
        exit
    }

}

if($Run){
     Invoke-Item $SelectedFile
}
else{
    $SelectedFile
}

