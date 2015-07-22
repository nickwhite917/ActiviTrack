#######################################################################
# Project: ActiviTrack Project Time Tracking Tool
# Project Date / Last Updated: 07/13/2015
# Script by: Nicholas White (517066) nick.white@parker.com
<# Script Version #> $sVer = 1.0
#######################################################################

#Function region
#region
function New-Project(){
  param ($ProjectName, $ProjectNumber)

  $obj = new-object PSObject
  $obj | add-member -type NoteProperty -Name ProjectName   -Value $ProjectName
  $obj | add-member -type NoteProperty -Name ProjectNumber  -Value $ProjectNumber

  return $obj
}

Function Status(){
    param ([string] $Information, [string] $Error, [string] $Update)

    $OutFilePath = "$strScriptPath\ActiviTrack_Log.act"

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
}

Function New-Command(){
  param ($CommandString)

  $CleanCommand = $CommandString.toUpper()
  
  $Commands = ("NEW","START","STOP","STATUS","COMMANDS","PRINT","BACK")

  If($Commands.Contains($CleanCommand)){
    $CommandType = $CleanCommand
  }
  Else{
    Return $CommandType = ""
  }

  $obj = new-object PSObject
  $obj | add-member -type NoteProperty -Name Command -Value $CommandType
  return $obj
}

Function Execute-Command($Command){

  switch ($Command) 
    { 
        "NEW" {Execute-New} 
        "START" {Execute-Start} 
        "STOP" {Execute-Stop} 
        "STATUS" {Execute-Status} 
        "COMMANDS" {Execute-Commands} 
        "PRINT" {Execute-Print} 
        "BACK" {Execute-Back}
    }
}
Function Execute-Status(){
    Space
    Write-Host "*********** Active Projects ***********"
    $Projects = Retrieve-Projects
    Foreach($Item in $Projects){
        $Name = $Item.ProjectName
        $Number = $Item.ProjectNumber
        Write-Host "$Number  |  $Name"
    }
    Space
    $Input = Read-Host "Which project would you like to know more about?"
        If($Input -eq "" -or $Input -eq $null -or $Input.Length -lt 2 -and (!(Project-Exists -ProjectName $Input) -or !(Project-Exists -ProjectNumber $Input))){
            Write-Host "Sorry, that does not exist in this list..."
            $InputNotValid = $True
        }
        else{
            $Project = Get-Project $Input
            $Name = $Project.ProjectName
            $Number = $Project.ProjectNumber
            Space
            Write-Host "$Number  |  $Name"
            Write-Host "***** Statistics *****"
        }

}
Function Space(){
    Write-Host ""
}
Function Get-Project($RawInput){
    
    $ProjectInput = $RawInput.toString()
    $Projects = Retrieve-Projects
    Foreach($Project in $Projects){
        $Name = $Project.ProjectName.toString().toUpper()
        $Number = $Project.ProjectNumber.toString()
        If($Name -eq $ProjectInput -or $Number -eq $ProjectInput){
            return $Project
        }
    }
}

Function Project-Exists($ProjectName, $ProjectNumber){
    $Projects = Retrieve-Projects
    Foreach($Project in $Projects){
        If($Project.ProjectName.toUpper() -eq $ProjectName.toUpper()){
            return $True
        }
        If($Project.ProjectNumber.toString() -eq $ProjectNumber.toString()){
            return $True
        }
    }
    Return $false
}

$ProjectFilePath = "$strScriptPath\Data\Projects.act"
    $TimeEntriesFilePath = "$strScriptPath\Data\TimeEntries.act"
    $StateFilePath = "$strScriptPath\Data\State.act"


Function Execute-New(){
    $InputNotValid = $True
    While($InputNotValid){
        $Input = Read-Host "Please enter the name of your new project: "
        If($Input -eq "" -or $Input -eq $null -or $Input.Length -lt 2){
            Write-Host "Sorry, that project name is not allowed..."
            $InputNotValid = $True
        }
        elseif(Project-Exists -ProjectName $Input){
            Write-Host "Sorry, that project name is already being used..."
            $InputNotValid = $True
        }
        Else{
            $ProjectName = $Input 
            $InputNotValid = $false
        }
    }
    $InputNotValid = $True
    While($InputNotValid){
        $Input = Read-Host "Please enter the project number: "
        If($Input -eq "" -or $Input -eq $null){
            $InputNotValid = $True
        }
        Else{
            $ProjectNumber = $Input 
            $InputNotValid = $false
        }
    }

    $Project = New-Project -ProjectName $ProjectName -ProjectNumber $ProjectNumber
    $CsvLine = New-CsvLine -Project $Project
    $CsvLine | Out-File -FilePath "$ProjectFilePath" -Append
    Write-Host "Project $ProjectName Added!"

}

function New-CsvLine($Project, $Statistic, $TimeEntry){
    If($Project -ne $null){
        $CsvLine += $Project.ProjectNumber
        $CsvLine += ','
        $CsvLine += $Project.ProjectName
    }
    ElseIf($Statistic -ne $null){
        $CsvLine  = $Statistic.Project
        $CsvLine += ','
        $CsvLine += $Statistic.TotalTime
        $CsvLine += ','
        $CsvLine += $Statistic.TimeThisWeek
        $CsvLine += ','
        $CsvLine += $Statistic.TimeToday
        $CsvLine += ','
        $CsvLine += $Statistic.AverageTimePerDay
    }
    ElseIf($TimeEntry -ne $null){
        $CsvLine  = $TimeEntry.Project
        $CsvLine += ','
        $CsvLine += $TimeEntry.TimeBegan
        $CsvLine += ','
        $CsvLine += $TimeEntry.TimeStopped
        $CsvLine += ','
        $CsvLine += $TimeEntry.TimeElapsed
    }

  return $CsvLine
}
Function Retrieve-Projects(){
   $Projects = @()
   $Content = Get-Content $ProjectFilePath
   Foreach($Item in $Content){
       If($Item.toUpper().Contains("PROJECT FILE")){
         Continue
       }
       If($Item.length  -lt 2){
         Continue
       }
       $Project = Parse-Project($Item)
       $Projects += $Project
   }
   Return $Projects
}
Function Parse-Project($Item){
   $Split = $Item.Split(',')
   $Project = New-Project -ProjectName $Split[1] -ProjectNumber $Split[0]
   Return $Project
}



#endregion


#Startup region
#region
#Try to exit out of any existing Word, or Excel instances

#exitConflictingPrograms
$Time =  Get-Date

Try{
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
}
Catch{}

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


#Create and/or Load Projects
#And previous time data

#Commands
# -New-Project
# -Started Working
# -Stopped Working
# -Status -Current working projects, Time summary
# -Com -Gives commands possible
# -Print

$ProjectFilePath = "$strScriptPath\Data\Projects.act"
$TimeEntriesFilePath = "$strScriptPath\Data\TimeEntries.act"
$StateFilePath = "$strScriptPath\Data\State.act"

If (!(Test-Path $ProjectFilePath)){
   "Project File" | Out-File $ProjectFilePath -Append
}
If (!(Test-Path $TimeEntriesFilePath)){
   "Time Entries File" | Out-File $TimeEntriesFilePath -Append
}
If (!(Test-Path $StateFilePath)){
   "Current State File" | Out-File $StateFilePath -Append
}

#Main Program Loop
$Exit = $false
While($Exit -ne $True){
    
    $Input = Read-Host "$~ "
    If($Input -eq "" -or $Input -eq $null){
        $Exit = $True
        Break
    }
    Else{
        $CommandObject = New-Command -CommandString $Input
        Execute-Command $CommandObject.Command
    }

}