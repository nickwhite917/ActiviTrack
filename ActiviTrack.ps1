#######################################################################
# Project: ActiviTrack Project Time Tracking Tool
# Project Date / Last Updated: 07/13/2015
# Script by: Nicholas White (517066) nick.white@parker.com
<# Script Version #> $sVer = 1.0
#######################################################################

#Function region
#region
function New-Project(){
  param ($ProjectName, $ProjectNumber, $ProjectLead, $TechnicalLead, $Programmer)

  $guid = [guid]::NewGuid()
  $guid = $guid.ToString()

  $obj = new-object PSObject
  $obj | add-member -type NoteProperty -Name ProjectGuid   -Value $guid
  $obj | add-member -type NoteProperty -Name ProjectName   -Value $ProjectName
  $obj | add-member -type NoteProperty -Name ProjectNumber  -Value $ProjectNumber
  $obj | add-member -type NoteProperty -Name ProjectLead   -Value $ProjectLead
  $obj | add-member -type NoteProperty -Name TechnicalLead -Value $TechnicalLead
  $obj | add-member -type NoteProperty -Name Programmer    -Value $Programmer

  return $obj
}

Function Status(){
param ([string] $Information, [string] $Error, [string] $Update)

$OutFilePath = "$strScriptPath\Checklist_Output.txt"

If($Information.Length -gt 1){
    Write-Host "Information: $Information"
    "Information: $Information" | Out-File $OutFilePath -Append
}

If($Error.Length -gt 1){
    Write-Host "Error: $Error"
    "Error: $Error" | Out-File $OutFilePath -Append
}

If($Update.Length -gt 1){
    Write-Host "Update: $Update"
    "Update: $Update" | Out-File $OutFilePath -Append
}
    
#Getting items from specified path
#Filtering for Checklists
#Information
#Processing Checklist
#Getting Checklist Info
#Generating CSV String

}
<#
.SYNOPSIS
Gets the hash value of a file or string
 
.DESCRIPTION
Gets the hash value of a file or string
It uses System.Security.Cryptography.HashAlgorithm (http://msdn.microsoft.com/en-us/library/system.security.cryptography.hashalgorithm.aspx)
and FileStream Class (http://msdn.microsoft.com/en-us/library/system.io.filestream.aspx)
Based on: http://blog.brianhartsock.com/2008/12/13/using-powershell-for-md5-checksums/ and some ideas on Microsoft Online Help
 
Be aware, to avoid confusions, that if you use the pipeline, the behaviour is the same as using -Text, not -File
 
.PARAMETER File
File to get the hash from.
 
.PARAMETER Text
Text string to get the hash from
 
.PARAMETER Algorithm
Type of hash algorithm to use. Default is SHA1
 
.EXAMPLE
C:\PS> Get-Hash "hello_world.txt"
Gets the SHA1 from myFile.txt file. When there's no explicit parameter, it uses -File
 
.EXAMPLE
Get-Hash -File "C:\temp\hello_world.txt"
Gets the SHA1 from myFile.txt file
 
.EXAMPLE
C:\PS> Get-Hash -Algorithm "MD5" -Text "Hello Wold!"
Gets the MD5 from a string
 
.EXAMPLE
C:\PS> "Hello Wold!" | Get-Hash
We can pass a string throught the pipeline
 
.EXAMPLE
Get-Content "c:\temp\hello_world.txt" | Get-Hash
It gets the string from Get-Content
 
.EXAMPLE
Get-ChildItem "C:\temp\*.txt" | %{ Write-Output "File: $($_)   has this hash: $(Get-Hash $_)" }
This is a more complex example gets the hash of all "*.tmp" files
 
.NOTES
DBA daily stuff (http://dbadailystuff.com) by Josep Martínez Vilà
Licensed under a Creative Commons Attribution 3.0 Unported License
 
.LINK
Original post: http://dbadailystuff.com/2013/03/11/get-hash-a-powershell-hash-function/
#>
function Get-Hash
{
     Param
     (
          [parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="set1")]
          [String]
          $text,
          [parameter(Position=0, Mandatory=$true, ValueFromPipeline=$false, ParameterSetName="set2")]
          [String]
          $file = "",
          [parameter(Mandatory=$false, ValueFromPipeline=$false)]
          [ValidateSet("MD5", "SHA", "SHA1", "SHA-256", "SHA-384", "SHA-512")]
          [String]
          $algorithm = "SHA1"
     )
     Begin
     {
          $hashAlgorithm = [System.Security.Cryptography.HashAlgorithm]::Create($algorithm)
     }
     Process
     {
          $md5StringBuilder = New-Object System.Text.StringBuilder 50
          $ue = New-Object System.Text.UTF8Encoding
 
          if ($file){
               try {
                    if (!(Test-Path -literalpath $file)){
                         throw "Test-Path returned false."
                    }
               }
               catch {
                    throw "Get-Hash - File not found or without permisions: [$file]. $_"
               }
               try {
                    [System.IO.FileStream]$fileStream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open);
                    $hashAlgorithm.ComputeHash($fileStream) | % { [void] $md5StringBuilder.Append($_.ToString("x2")) }
               }
               catch {
                    throw "Get-Hash - Error reading or hashing the file: [$file]"
               }
               finally {
                    $fileStream.Close()
                    $fileStream.Dispose()
               }
          }
          else {
               $hashAlgorithm.ComputeHash($ue.GetBytes($text)) | % { [void] $md5StringBuilder.Append($_.ToString("x2")) }
          }
 
          return $md5StringBuilder.ToString()
     }
}
Function MakeLocal($Checklist){
    $Name = $Checklist
    Copy-Item -Path $Checklist.FullName -Destination "$strScriptPath\$Name"
    $Checklist = Get-Item -Path "$strScriptPath\$Name"
    Return $Checklist
}
Function RemoveLocal($Checklist){

    Remove-Item -Path $Checklist.FullName -Force

}
#endregion


#Startup region
#region
#Try to exit out of any existing Word, or Excel instances
exitConflictingPrograms
$Time =  Get-Date

$pshost = get-host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 150
$pswindow.buffersize = $newsize

$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 150
$pswindow.windowsize = $newsize


#Resolve script/executable path
try{
		# Find the program's location [Depends on if it is running as a .ps1 or .exe]
		$strScriptPath = Split-Path (Resolve-Path $myInvocation.MyCommand.Path)
		if ($strScriptPath -match 'Quest Software') {$strScriptPath = [System.AppDomain]::CurrentDomain.BaseDirectory}
}
catch{
	Break;
}

if($strScriptPath -eq $null -or $strScriptPath -eq ""){
	Break;
}

If(Test-Path "$strScriptPath\Checklist_Data.txt"){
    $Prefix = (((Get-Date -Format o).Replace("T","-")).Replace(":","")).Substring(0,17) #Date Prefix Formatter
    Rename-Item "$strScriptPath\Checklist_Data.txt" -NewName "$Prefix Checklist_Data.txt"
}
If(Test-Path "$strScriptPath\Checklist_Output.txt"){
    $Prefix = (((Get-Date -Format o).Replace("T","-")).Replace(":","")).Substring(0,17) #Date Prefix Formatter
    Rename-Item "$strScriptPath\Checklist_Output.txt" -NewName "$Prefix Checklist_Output.txt"
}
If(Test-Path "$strScriptPath\Checklist_Progress.txt"){
    $Prefix = (((Get-Date -Format o).Replace("T","-")).Replace(":","")).Substring(0,17) #Date Prefix Formatter
    Rename-Item "$strScriptPath\Checklist_Progress.txt" -NewName "$Prefix Checklist_Progress.txt"
}

Status -Information "Script is being executed from  : $strScriptPath."
Status -Information "Checklist_Data.txt   located at: $strScriptPath\Checklist_Data.txt"
Status -Information "Checklist_Output.txt located at: $strScriptPath\Checklist_Output.txt."

#Create and/or Load Projects
#And previous time data

#Commands
# -New-Project
# -Started Working
# -Stopped Working
# -Status -Current working projects, Time summary
# -Com -Gives commands possible
# -Print
# -



#Main Program Loop
While($Exit -ne $True){
    
    $Command = Read-Host "$~ "
    If($Command -eq "" -or $Command -eq $null){
        Break;
    }
    Else{
        
    }

}

$DirtyPath = Read-Host "Please enter the path of the SharePoint site"

$Path = Sanitize($DirtyPath)

If($Path -eq "NULL"){
    Status -Error "The path to $DirtyPath was not found, or contains an error. Stopping execution..."
    Break;
}

Status -Information "***| Beginning execution. Console output will be displayed here."
Status -Information "***| Start Time: $Time"
Status -Information "***| SharePoint Target: $Path"

#$Path = "//polteams/sites/cor/MSS/MS SAR 250K299K"

Status -Update "Fetching items from SharePoint..."
$ItemsUnderPath = Get-ChildItem $Path -Force -Recurse
Status -Update "Items found!"

$Count = $ItemsUnderPath.Count
Status -Information "Specified path contains $Count items."

$Folders = @()
$Files = @()
$Checklists = @()

Status -Update "Creating new instance of Excel..."
$xl = NewExcel;
Status -Update "New instance of Excel created!"

$ChecklistCount = 0
$FolderCount = 0
$FileCount = 0

Foreach($Item in $ItemsUnderPath){
    if($Item -is [System.IO.DirectoryInfo]){
        $Folders += $Item
        $FolderCount = $Folders.Count
    }
    else{
        $Files += $Item
        $FileCount = $Files.Count
        if($Item.Name.ToUpper().Contains("CHECKLIST")){
            $Checklists += $Item
            $ChecklistCount = $Checklists.Count
        }
        elseif(isChecklist -Item $Item -Excel $xl){
            $Checklists += $Item
            $ChecklistCount = $Checklists.Count
        }
    }
    Write-Host "`rStatus: $FolderCount Folders | $FileCount Files | $ChecklistCount Checklists found. Searching..." -NoNewline
}
Write-Host ""
Status -Information "Search complete! $FolderCount Folders | $FileCount Files | $ChecklistCount Checklists found."

$Count = $Folders.Count
Status -Information "Specified path contains $Count folders."
$Count = $Files.Count
Status -Information "Specified path contains $Count files."
$Count = $Checklists.Count
Status -Information "Specified path contains $Count Development Checklists."

#Send Field Names to CSV
$CsvLine = New-CsvLine -Filter "NONE" -Project "NONE" -Progress "NONE"   
$CsvLine | Out-File -FilePath "$strScriptPath\Checklist_Data.txt" -Append


Foreach($Checklist in $Checklists){
    Write-Host " "
    Write-Host " "
    Status -Update "Making a local copy of the checklist..."
    $Checklist = MakeLocal($Checklist)
    Status -Update "Processing $Checklist checklist..."
    Status -Update "Opening checklist..."
    $wb = $xl.Workbooks.open($Checklist.FullName)
    
    $ws = $wb.WorkSheets.item("Development Checklist") 
    $ws.Activate()

    Status -Update "Gathering information about the checklist's type..."
    $ConfigSettings      = Get_OleObjectValues($ws)

    $FilterNonMainFrame  = $ConfigSettings.FilterNonMainFrame
    $FilterPackage       = $ConfigSettings.FilterPackage
    $FilterConfiguration = $ConfigSettings.FilterConfiguration
    $FilterMainframe     = $ConfigSettings.FilterMainframe
    $FilterDevelopment   = $ConfigSettings.FilterDevelopment

    $Filter = New-Filter -FilterNonMainFrame $FilterNonMainFrame -FilterPackage $FilterPackage -FilterConfiguration $FilterConfiguration -FilterMainframe $FilterMainframe -FilterDevelopment $FilterDevelopment
    
    $ProjectName = $ws.Range("I1").Text
    $ProjectNumber = $ws.Range("I2").Text
    $ProjectLead = $ws.Range("I3").Text
    $TechnicalLead = $ws.Range("L2").Text
    $Programmer = $ws.Range("L3").Text

    $Project = New-Project -ProjectName $ProjectName -ProjectNumber $ProjectNumber -ProjectLead $ProjectLead -TechnicalLead $TechnicalLead -Programmer $Programmer
    
    $Progress = @()

    Status -Update "Processing step completion data..."
    for($i = 15; $i -lt 100; $i++){
        $CurrentStep = $i - 15
        $Percent = ($CurrentStep / 85) * 100
        $Percent = "{0:N2}" -f $Percent
        Write-Host "`rStatus: Processing step # $CurrentStep of 85 | $Percent % Complete" -NoNewline
        $Phase    = $ws.Range("A$I").Text
        $Step     = $ws.Range("B$I").Text
        $Status   = $ws.Range("J$I").Text

        if($ws.Range("A$I").Height -gt 1){
            $LineItem = New-Step -Phase $Phase -Step $Step -Status $Status -ProjectGuid $Project.ProjectGuid
            $Progress += $LineItem
        }
    }

    Status -Information "Creating line in csv file for $Checklist checklist..."
    $CsvLine = New-CsvLine -Filter $Filter -Project $Project -Progress $Progress -ProjectGuid $ProjectGuid
    $CsvLine | Out-File -FilePath "$strScriptPath\Checklist_Data.txt" -Append
    ProgressDataOut -Progress $Progress
    #Write to file

    $wb.Close($False)

    RemoveLocal($Checklist)

    Status -Update "Checklist $checklist complete"
    $ChecklistsCompleted ++

    $ChecklistProgress = 100 * ($ChecklistsCompleted/$ChecklistCount)
    $ChecklistProgress = "{0:N2}" -f $ChecklistProgress
    Status -Information "Checklist Progress: $CheckListProgress % Complete"
}
$ChecklistProgress = 100 * ($ChecklistsCompleted/$ChecklistCount)
$ChecklistProgress = "{0:N2}" -f $ChecklistProgress
Status -Information "Checklist Progress: $CheckListProgress % Complete"
