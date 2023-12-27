<#
    Program Name : Pelican-Basic-FM
    Verison : 1.3.0
    Description : Basic script for Model Federation automation and will be improved occasionally.
    
    New in v1.3.0:
    - Copy only updated NWCs
    - Update progress reporting (minor)
#>

#Read config files for file and folder locations
$ConfigLocation = ([xml](Get-Content Pelican-Basic-FM_Config.xml)).Config

$BatchUtilityProcess = $ConfigLocation.BatchUtility
$LogFolder = $ConfigLocation.LogFolder
$BackupDirectory = $ConfigLocation.BackupLocation
$MainBuildFolder = $ConfigLocation.MainBuildFolder

$ByLevel = $ConfigLocation.BatchTextFile.ByLevel
$ByBuilding = $ConfigLocation.BatchTextFile.ByBuilding
$ByOverall = $ConfigLocation.BatchTextFile.ByOverall
$ByFederatedModel = $ConfigLocation.BatchTextFile.ByFederatedModel
$ByFinalFM = $ConfigLocation.BatchTextFile.ByFinalFM

$ByLevelOut = $ConfigLocation.TempBuildFolder.ByLevel
$ByBuildingOut = $ConfigLocation.TempBuildFolder.ByBuilding
$ByOverallOut = $ConfigLocation.TempBuildFolder.ByOverall
$ByFederatedModelOut = $ConfigLocation.TempBuildFolder.ByFederatedModel
$ByFinalFMOut = $ConfigLocation.TempBuildFolder.ByFinalFM

$TempNWC_All = $ConfigLocation.TempNWCFolder.NWC_ALL
$TempNWC_CM = $ConfigLocation.TempNWCFolder.NWC_CM
$TempNWC_DM = $ConfigLocation.TempNWCFolder.NWC_DM
$TempNWC_EM = $ConfigLocation.TempNWCFolder.NWC_EM

$MainNWC_All = $ConfigLocation.MainNWCFolder.NWC_ALL
$MainNWC_CM = $ConfigLocation.MainNWCFolder.NWC_CM
$MainNWC_DM = $ConfigLocation.MainNWCFolder.NWC_DM
$MainNWC_EM = $ConfigLocation.MainNWCFolder.NWC_EM

$ArgumentsByLevel = '/i "{0}" /od "{1}"' -f $ByLevel, $ByLevelOut
$ArgumentsByBuilding = '/i "{0}" /od "{1}"' -f $ByBuilding, $ByBuildingOut
$ArgumentsByOverall = '/i "{0}" /od "{1}"' -f $ByOverall, $ByOverallOut
$ArgumentsByFederatedModel = '/i "{0}" /od "{1}"' -f $ByFederatedModel, $ByFederatedModelOut
$ArgumentsByFinalFM = '/i "{0}" /od "{1}"' -f $ByFinalFM, $ByFinalFMOut

#Function for writing log file
function WriteLog
{
    Param ([string]$LogString)
    $LogFile = "$LogFolder\Pelican-Basic-MASTER-log.txt"
    $DateTime = "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $LogMessage = "$Datetime $LogString"
    Add-content $LogFile -value "$LogMessage"
}

WriteLog "---------------- START ----------------"
  
#Backup the last NWCs
try{
    $i = 0
    $Files = Get-ChildItem $TempNWC_All -Exclude "_Archived" | Get-ChildItem  -Recurse
    $FileDest = "$BackupDirectory\$((Get-Date).ToString('yyyy-MM-dd'))"
    ForEach($File in $Files){
        $i = $i+1
        $FFileName = "$FileDest\{0}" -f $File.Name
        if ($File.Name -like "*-CM*")
            {
                $FFileName = "$FileDest\CM-NWCS\{0}" -f $File.Name
            }
        if ($File.Name -like "*-DM*")
            {
                $FFileName = "$FileDest\DM-NWCS\{0}" -f $File.Name
            }
        if ($File.Name -like "*-EM*")
            {
                $FFileName = "$FileDest\EM-NWCS\{0}" -f $File.Name
            }
        New-Item -ItemType File -Path $FFileName -Force
        Copy-Item -Path $File.FullName -Destination $FFileName -Force
        Write-Progress -Activity ("Backing up previous NWC files {0}/{1} ({2})..." -f $i, $Files.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
        }
        WriteLog "[COMPLETED] Backup previous NWC files"
    }
catch{
    $BackupException = $_.Exception.Message
    Write-Output "$BackupException"
    WriteLog "[ERROR] Backup previous NWC files"
    WriteLog "[ERROR] $BackupException"
    }

#Copy the latest NWCs into the temporary build folder
try{
    $i = 0
    #Search only NWCs with the last modified date is a day before the script is run
    $Files = Get-ChildItem $MainNWC_All -Exclude "*_Archived*","*_Rejected*" | Get-ChildItem -Recurse -Filter "*.nwc" | Where-Object { $_.LastWriteTime -gt $(((Get-Date).AddDays(-1)).ToString('yyyy-MM-dd')) }
    ForEach($File in $Files){
        $i = $i+1
        $FFileName = "$TempNWC_All\{0}" -f $File.Name
        if ($File.Name -like "*-CM*")
            {
                $FFileName = "$TempNWC_All\CM-NWCS\{0}" -f $File.Name
            }
        if ($File.Name -like "*-DM*")
            {
                $FFileName = "$TempNWC_All\DM-NWCS\{0}" -f $File.Name
            }
        if ($File.Name -like "*-EM*")
            {
                $FFileName = "$TempNWC_All\EM-NWCS\{0}" -f $File.Name
            }
        New-Item -ItemType File -Path $FFileName -Force
        Copy-Item -Path $File.FullName -Destination $FFileName -Force
        Write-Progress -Activity ("Copy latest NWC files {0}/{1} ({2})..." -f $i, $Files.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
        }
    WriteLog "[COMPLETED] Copy Latest NWC files"
}
catch{
    $CopyNWCException = $_.Exception.Message
    Write-Output "$CopyNWCException"
    WriteLog "[ERROR] Copy Latest NWC files"
    WriteLog "[ERROR] $CopyNWCException"
}

#Start the Batch Utility process and build Model ByLevel > ByBuilding > ByOverall > Federated Model and save the model to temporary folder
try{
    Start-Process $BatchUtilityProcess -ArgumentList $ArgumentsByLevel -Wait -NoNewWindow
    WriteLog "[COMPLETED] Building model By Level"
    Start-Process $BatchUtilityProcess -ArgumentList $ArgumentsByBuilding -Wait -NoNewWindow
    WriteLog "[COMPLETED] Building model By Building"
    Start-Process $BatchUtilityProcess -ArgumentList $ArgumentsByOverall -Wait -NoNewWindow
    WriteLog "[COMPLETED] Building model By Overall"
    Start-Process $BatchUtilityProcess -ArgumentList $ArgumentsByFederatedModel -Wait -NoNewWindow
    WriteLog "[COMPLETED] Building model By Federated Model"
    Start-Process $BatchUtilityProcess -ArgumentList $ArgumentsByFinalFM -Wait -NoNewWindow
    WriteLog "[COMPLETED] Building model By Final Federated Model"
    }
catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    WriteLog "[ERROR] $BuildException"
    }

#Copy the latest model build to the NWD folder
try{
    #By Level Folder
    $ModelType = "CM", "DM", "EM"
    ForEach($Type in $ModelType) {
        $i = 0
        $Files = Get-ChildItem $ByLevelOut -Recurse -Filter ("*-{0}.nwd" -f $Type)
        ForEach($File in $Files){
            $i = $i+1
            $FileDestination = "$MainBuildFolder\{0}\By Level\{1}" -f $Type, $File.Name
            Write-Progress -Activity ("Copying latest Federated Model Files By Level ({0})..." -f $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
            New-Item -ItemType File -Path $FileDestination -Force
            Copy-Item -Path $File.FullName -Destination $FileDestination -Force
            }
        }

    #By Building Folder
    ForEach($Type in $ModelType) {
        $i = 0
        $Files = Get-ChildItem $ByBuildingOut -Recurse -Filter ("*-{0}.nwd" -f $Type)
        ForEach($File in $Files){
            $i = $i+1
            $FileDestination = "$MainBuildFolder\{0}\By Building\{1}" -f $Type, $File.Name
            Write-Progress -Activity ("Copying latest Federated Model Files By Building ({0})..." -f $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
            New-Item -ItemType File -Path $FileDestination -Force
            Copy-Item -Path $File.FullName -Destination $FileDestination -Force
            }
        }

    #By Overall Folder
    ForEach($Type in $ModelType) {
        $i = 0
        $Files = Get-ChildItem $ByOverallOut -Recurse -Filter ("*-{0}.nwd" -f $Type)
        ForEach($File in $Files){
            $i = $i+1
            $FileDestination = "$MainBuildFolder\{0}\By Building\{1}" -f $Type, $File.Name
            Write-Progress -Activity ("Copying latest Federated Model Files By Overall ({0})..." -f $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
            New-Item -ItemType File -Path $FileDestination -Force
            Copy-Item -Path $File.FullName -Destination $FileDestination -Force
            }
        }

    #By Federated Model Folder
    $i = 0
    $Files = Get-ChildItem $ByFinalFMOut -Recurse -Filter ("*-FM.nwd")
    ForEach($File in $Files){
        $i = $i+1
        $FileDestination = "$MainBuildFolder\FM\{0}" -f $File.Name
        Write-Progress -Activity ("Copying latest Federated Model Files FEDERATED MODEL ({0})..." -f $Files.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
        New-Item -ItemType File -Path $FileDestination -Force
        Copy-Item -Path $File.FullName -Destination $FileDestination -Force
        }
    WriteLog "[COMPLETED] Copy latest Federated Model files to main folder"
    }
catch{
    $CopyException = $_.Exception.Message
    Write-Output "$CopyException"
    WriteLog "[ERROR] Copy latest Federated Model files to main folder"
    WriteLog "[ERROR] $CopyException"
    }
WriteLog "---------------- END ----------------"