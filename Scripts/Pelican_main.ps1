<#
    Program Name : Pelican_main
    Version : 5.2.0
    Description : Build affected NWF if there is new or retired models. Build necessary Federated Models
    Author : Lawrenerno Jinkim (lawrenerno.jinkim@exyte.net)
#>

. .\Config.ps1

WriteLog-Full "Running build on local Computer: $env:computername"

#Backup the previous NWCs folder
try{
    WriteLog-Full "Backup previous NWCs folder: $TempNWC_All"
    $i = 0
    $Files = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Filter "*.nwc"  -Recurse
    $FileDest = "$BackupDirectory\$((Get-Date).ToString('yyyy-MM-dd'))"
    If(!(Test-Path -Path $FileDest)){
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
        }
    else{
        WriteLog-Full ("Backup folder already exist :{0}" -f $((Get-Date).ToString('yyyy-MM-dd')))
        }
    }
catch{
    $BackupException = $_.Exception.Message
    WriteLog-Full "$BackupException" -Type ERROR
    $BuildSuccess=$false
    }

#Copy latest NWCs into the temporary build folder
try{
    WriteLog-Full "Copy UPDATED and NEW NWCs into temporary build folder"
    $i = 0
    #Search only NWCs with the last modified date is a day before the script is run
    $Files = Get-ChildItem $MainNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc" | Where-Object { $_.LastWriteTime -gt $(((Get-Date).AddDays(-1)).ToString('yyyy-MM-dd')) }
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
}
catch{
    $CopyNWCException = $_.Exception.Message
    WriteLog-Full "$CopyNWCException" -Type ERROR
    $BuildSuccess=$false
}

#Move Rejected models
try{
    #Regex for correct file naming
    $NPattern = "^XPG\w{2,3}-\w{3,4}-\d{3}-.{2}-\w{5}-\w{3}-(CM|DM|EM)(\.|[-_]ROOM\.|[-_]MAXFIT\.|[-_]AMHS\.|-ST\.)nwc$"
    $i = 0
    $Files = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","_Retired","_New.txt" | Get-ChildItem  -Recurse -FIlter "*.nwc"
    New-Item -ItemType Directory -Path $RejectedFolder -Force
    $FList = $Files -notmatch $NPattern
    If($Flist){
        WriteLog-Full ("Moving {0} rejected model into: {1}" -f $Flist.count, $RejectedFolder)
        ForEach($File in $FList){
            $i = $i+1
            Write-Progress -Activity ("Moving rejected models {0}/{1} ({2})..." -f $i, $Flist.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Flist.count)*100)
            $FFileName = "$RejectedFolder\{0}" -f $File.Name
            New-Item -ItemType File -Path $FFileName -Force
            Move-Item -Path $File.FullName -Destination $FFileName -Force
            WriteLog-Full ("Rejected model: {0}" -f $File.Name)
                }
        }
    }
catch{
    $Exception = $_.Exception.Message
    WriteLog-Full "$Exception" -Type ERROR
    $BuildSuccess=$false
    }

################ NWF BUILD REGION START ################

$F26FMlist = "F26_APB1-FM","F26_APB2-FM","F26_FAB-FM","F26_BCS-FM","LK1-FM"
$ModelArrayLevel = "F26_APB1","F26_APB2","F26_FAB","PGB","PGP"
$ModelArrayAncillaryBuilding = "F26_BCS","BG1","BG2","LB1","LK1","P09","P12","PGC","WTY"
$ModelPhase = "_DM","_CM","_EM"
$WithRetiredModel = $true
$WithNewModel = $true
$WithUpdatedModel = $true
$NWCList = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc"
$NWCList_today = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc" | Where-Object { $_.LastWriteTime -gt $(((Get-Date).AddDays(-1)).ToString('yyyy-MM-dd')) }

#Get list of model to be rebuild
try{
    #Check if there is modified nwc files
    If($NWCList_today){
        #Get the list of model that need to be rebuild for by level
        $ModeltoRebuildByLevel = @()
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayLevel){
                $Array = (Get-Variable $Model$Phase).Value
                $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $NWCList_today -Stage Level
                $ModeltoRebuildByLevel += $ModelList
            }
        }
        #Get the list of model that need to be rebuild for main building
        $ModeltoRebuildByBuildingMain = @()
        ForEach($a in $ModelPhase){
            ForEach($b in $ModelArrayLevel){
                $fileN = ("$b*{0}" -f $a.Replace("_","-"))
                If($ModeltoRebuildByLevel -like $fileN){
                    $ModeltoRebuildByBuildingMain += ("$b{0}" -f $a.Replace("_","-"))
                }
        
            }
        }
        #Get the list of model that need to be rebuild for ancillary building
        $ModeltoRebuildByBuilding = @()
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayAncillaryBuilding){
                $Array = (Get-Variable $Model$Phase).Value
                $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $NWCList_today -Stage Building
                $ModeltoRebuildByBuilding += $ModelList
            }
        }
        If(!($ModeltoRebuildByBuilding)){
            WriteLog-Full "All Ancillary building is up to date." -Type INFO
            }
        #Combine list of building model into one list
        $ModeltoRebuildByBuilding = $($ModeltoRebuildByBuildingMain;$ModeltoRebuildByBuilding)

        #List of FM model to be rebuild
        $ModeltoRebuildByFM = ($ModeltoRebuildByBuilding -replace "[CDE]M","FM") | Sort -Unique

        #Get the list of PG model to be rebuild
        $ModeltoRebuildByOverall = @()
        ForEach($Phase in $ModelPhase){
            $Phase = $Phase.Replace("_","-")
            If($ModeltoRebuildByBuilding -like "*$Phase"){
                $ModeltoRebuildByOverall += "PG{0}" -f $Phase
            }
        }

        #Final FM model to be rebuild
        $ModeltoRebuildByFinal = @()
        If($ModeltoRebuildByFM){
            $ModeltoRebuildByFinal += "PG-FM"
            }
        ForEach($alpha in $F26FMlist){
            If($ModeltoRebuildByFM -contains $alpha){
                $ModeltoRebuildByFinal += "F26-FM"
                Break
                }
        }

        #Temp output
        Write-Output "By Level: "
        Write-Output $ModeltoRebuildByLevel | Sort
        Write-Output "`nBy Building: "
        Write-Output $ModeltoRebuildByBuilding | Sort
        Write-Output "`nBy FM: "
        Write-Output $ModeltoRebuildByFM
        Write-Output "`nBy Overall: "
        Write-Output $ModeltoRebuildByOverall
        Write-Output "`nBy Final FM: "
        Write-Output $ModeltoRebuildByFinal
    }
    else{
        $WithUpdatedModel = $false
        }
}

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess = $false
    }

If($WithUpdatedModel){
    $ModeltoRebuildByLevel = $ModeltoRebuildByLevel | ForEach-Object {"$_.nwf"}
    $ModeltoRebuildByBuilding = $ModeltoRebuildByBuilding | ForEach-Object {"$_.nwf"}
    $ModeltoRebuildByFM = $ModeltoRebuildByFM | ForEach-Object {"$_.nwf"}
    $ModeltoRebuildByOverall = $ModeltoRebuildByOverall | ForEach-Object {"$_.nwf"}
    $ModeltoRebuildByFinal = $ModeltoRebuildByFinal | ForEach-Object {"$_.nwf"}
    }
else{
    $WithUpdatedModel = $false
    }

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwd"
$Buildinglist = "BG1-","BG2-","F26_APB1-","F26_APB2-","F26_FAB-","F26_BCS-","LB1-","LK1-","P09-","P12-","PGB-","PGP-","PGC-","WTY-"

<#
#Check for retired model
If(!(Test-Path "$TempNWC_All\_Retired")){
        WriteLog-Full "Retired Model folder does not exist: $TempNWC_All\_Retired - Building model without retired model list" -Type WARN
        $WithRetiredModel = $false
    }
else{
    $NWCList_retired = Get-ChildItem "$TempNWC_All\_Retired" -Filter "*.nwc"
    If(!(Get-ChildItem "$TempNWC_All\_Retired" -Filter "*.nwc")){
        WriteLog-Full "No retired model found - Building model without retired model list" -Type INFO
        $WithRetiredModel = $false
    }
    else{
        WriteLog-Full ("{0} Retired model found - Building model with retired model list" -f $NWCList_Retired.count) -Type INFO
        ForEach($r in $NWCList_Retired){
            WriteLog-Full ("Retired model: {0}" -f $r.Name)
            }
        }
}#>

#Check for retired model
If(!(ReadExcelFile -Path $NRFile -SheetName $RSheet -Mode Read)){
    WriteLog-Full ("There is no retired model file found in: {0}" -f (Split-Path $NRFile -Leaf)) -Type INFO
    $WithRetiredModel = $false
    }
else{
    $ListOfRetiredModel = ReadExcelFile -Path $NRFile -SheetName $RSheet -Mode Read
    WriteLog-Full ("{0} Retired model found - Building model with retired model list" -f $ListOfRetiredModel.count) -Type INFO
    $ListOfRetiredModel | ForEach-Object {WriteLog-Full "Retired model: $_"}
    }

#Check if there is new model
If(!(ReadExcelFile -Path $NRFile -SheetName $NSheet -Mode Read)){
    WriteLog-Full ("There is no new model file found in: {0}" -f (Split-Path $NRFile -Leaf)) -Type INFO
    $WithNewModel = $false
    }
else{
    $ListofNewModel = ReadExcelFile -Path $NRFile -SheetName $NSheet -Mode Read
    WriteLog-Full ("{0} New model found - Building model with new model list" -f $ListofNewModel.count) -Type INFO
    $ListofNewModel | ForEach-Object {WriteLog-Full "New model: $_"}
    }

#Rebuild NWF model only if there is new or retired model
try{
    #$ListofNewModel = Get-Content $NewModelFile
    If(!($WithRetiredModel) -and !($WithNewModel)){
        WriteLog-Full "Skip rebuild NWF files. All NWF is up to date." -Type INFO
        }
    else{
        WriteLog-Full "Rebuilding nwf files..."
        $ListOfNewAndRetiredModel = @()
        If($WithNewModel){
            $ListOfNewAndRetiredModel = $ListofNewModel
            }
        If($WithRetiredModel){
            #$NWCList_retired = Get-ChildItem "$TempNWC_All\_Retired" -Filter "*.nwc"
            #BEFORE FIX
            #$ListOfNewAndRetiredModel += $NWCList_retired.Name
            #$ListOfRetiredModel = $NWCList_retired | ForEach-Object {$_.Name}
            $ListOfNewAndRetiredModel = $($ListofNewModel;$ListOfRetiredModel)
            }

        #$ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired" -Include $ListOfNewAndRetiredModel | Get-ChildItem -Recurse -Filter "*.nwc"
        #$ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_New.txt" -Include $ListOfNewAndRetiredModel -Recurse | Where-Object {$_.FullName -notlike "*\_*"}
        #$ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_New.txt" -Include $ListOfNewAndRetiredModel -Recurse | Where-Object {$_.FullName -notlike "*\_Archived*"}
        
        #Build NWF By Level model for DM CM EM
        $NWFModeltoRebuildByLevel = @()
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayLevel){
                $Array = (Get-Variable $Model$Phase).Value
                RebuildNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $ListOfNewAndRetiredModel -OutFolder $NWFFolderByLevel -Stage Level
                $ModelList = GetModelToRebuild_List -ModelArray $Array -FileListT $ListOfNewAndRetiredModel -Stage Level
                $NWFModeltoRebuildByLevel += $ModelList
            }
        }

        #Build NWF for Main Building model for DM CM EM
        $NWFList_Temp = Get-ChildItem $NWFFolderByLevel -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
        $NWFtoRebuildByBuildingMain = @()
        If($NWFList_Temp){
            ForEach($a in $ModelPhase){
                ForEach($b in $ModelArrayLevel){
                    $Array = (Get-Variable $b$a).Value
                    $fileN = ("$b*{0}" -f $a.Replace("_","-"))
                    If($NWFModeltoRebuildByLevel -like $fileN){
                        $NWFtoRebuildByBuildingMain += ("$b{0}" -f $a.Replace("_","-"))
                    }
                    RebuildNWF-Dynamic -ModelArray $Array -FileList $NWDList -FileListT $NWFList_Temp -OutFolder $NWFFolderByBuilding -Stage Building -BuildingType Main
        
                }
            }
        }

        #Build NWF for Ancillary By Building model for DM CM EM
        Write-Output "DEBUG--"
        Write-Output $ListOfNewAndRetiredModel
        Write-Output "--DEBUG"
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayAncillaryBuilding){
                $Array = (Get-Variable $Model$Phase).Value
                RebuildNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $ListOfNewAndRetiredModel -OutFolder $NWFFolderByBuilding -Stage Building -BuildingType Ancillary
                $ModelList = GetModelToRebuild_List -ModelArray $Array -FileListT $ListOfNewAndRetiredModel -Stage Building
                $ToRebuildByBuilding += $ModelList
            }
        }

        $ToRebuildByBuildingMainFM = ($NWFtoRebuildByBuildingMain -replace "[CDE]M","FM") | Sort -Unique
        $ToRebuildByBuildingFM = ($ToRebuildByBuilding -replace "[CDE]M","FM") | Sort -Unique
        $FMModelToBeRebuild = $($ToRebuildByBuildingMainFM;$ToRebuildByBuildingFM)

        #Check if need to rebuild FM NWF files
        If($ListOfNewAndRetiredModel -like "*-EM*.nwc"){
            If($ListOfNewAndRetiredModel.count -le 1){
                WriteLog-Full ("EM model found: {0}" -f $ListOfNewAndRetiredModel.Name) -Type INFO
            }
            else{
                WriteLog-Full ("{0} EM model found:" -f ($ListOfNewAndRetiredModel -like "*-EM*.nwc").count) -Type INFO
                $EMModelList = ($ListOfNewAndRetiredModel -like "*-EM*.nwc") | ForEach-Object { "$_"}
                $EMModelList | ForEach-Object {WriteLog-Full $_}
                #Write-Output "DEBUG--"
                #Write-Output $EMModelList
                #Write-Output "--DEBUG"
                #WriteLog-Full ("EM model found: {0}" -f ($ListOFNWC.Name -like "*-EM*.nwc")) -Type INFO
                }
            $ListOFNWCEM = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_New.txt" -Include $EMModelList -Recurse | Where-Object {$_.FullName -notlike "*\_*"}
            $FMModeltoRebuildByLevel = @()
            ForEach($Phase in $ModelPhase){
                ForEach($Model in $ModelArrayLevel){
                    $Array = (Get-Variable $Model$Phase).Value
                    $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $ListOFNWCEM -Stage Level
                    $FMModeltoRebuildByLevel += $ModelList
                }
            }

            $FMModeltoRebuildByBuildingMain = @()
            $FMModeltoRebuildByBuildingMain = ForEach($a in $ModelPhase){
                ForEach($b in $ModelArrayLevel){
                    $fileN = ("$b*{0}" -f $a.Replace("_","-"))
                    If($FMModeltoRebuildByLevel -like $fileN){
                        #$FMModeltoRebuildByBuildingMain += ("$b{0}" -f $a.Replace("_","-"))
                        "$b{0}" -f $a.Replace("_","-")
                    }
                }
            }
            #Get the list of model that need to be rebuild for ancillary building
            $FMModeltoRebuildByBuilding = @()
            ForEach($Phase in $ModelPhase){
                ForEach($Model in $ModelArrayAncillaryBuilding){
                    $Array = (Get-Variable $Model$Phase).Value
                    $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $ListOFNWCEM -Stage Building
                    $FMModeltoRebuildByBuilding += $ModelList
                }
            }

            #Rebuild PG-EM
            $NWFList_Temp = Get-ChildItem $NWFFolderByBuilding -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
            RebuildPGEM -BuildingList $Buildinglist -FileList $NWDList -FileListT $NWFList_Temp -OutFolder $NWFFolderByOverall

            #Combine list of building model into one list
            $FMModeltoRebuildByBuilding = $($FMModeltoRebuildByBuildingMain; $FMModeltoRebuildByBuilding)
            
            #List of FM model to be rebuild
            $FMModeltoRebuildByFM = ($FMModeltoRebuildByBuilding -replace "[CDE]M","FM") | Sort -Unique
            WriteLog-Full "Building NWF FM Model for:"
            $FMModeltoRebuildByFM | ForEach-Object {WriteLog-Full "$_.nwf"}
            $FMModeltoInclude = @()
            $FMModelToBeRebuild = $($FMModeltoRebuildByFM;$FMModelToBeRebuild)
            $FMModeltoInclude = $FMModelToBeRebuild | ForEach-Object {"$_.nwf"}
            RebuildNWFFM -ModelArray $FMModeltoRebuildByFM -FileList $NWDList -OutFolder $NWFFolderByFM
            }
        }
}

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess = $false
    }

$FinalFMFilter = "PG-FM.nwf","F26-FM.nwf"
$NWFList_today = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_ByLevel = Get-ChildItem $NWFFolderByLevel -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_ByBuilding = Get-ChildItem $NWFFolderByBuilding -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_ByOverall = Get-ChildItem $NWFFolderByOverall -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_ByFM = Get-ChildItem $NWFFolderByFM -Recurse -Filter "*.nwf" -Include $FMModeltoInclude | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_ByFM_All = Get-ChildItem $NWFFolderByFM -Recurse -Filter "*.nwf" -Include $FinalFMFilter

#Updating the list of models need to be rebuild (new, retired and modified)
If($NWFList_today -or $WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
    If($NWFList_ByLevel-or $WithUpdatedModel){
        $ModeltoRebuildByLevelNWF = $NWFList_ByLevel | ForEach-Object { $_.Name }
        $ModeltoRebuildByLevelNWF = $($ModeltoRebuildByLevel;$ModeltoRebuildByLevelNWF)
        $ToIncludeByLevel = $ModeltoRebuildByLevelNWF | Sort -Unique
        }
    If($NWFList_ByBuilding -or $WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
        $ModeltoRebuildByBuildingNWF = $NWFList_ByBuilding | ForEach-Object { $_.Name }
        $ModeltoRebuildByBuildingNWF = $($ModeltoRebuildByBuilding;$ModeltoRebuildByBuildingNWF)
        $ToIncludeByBuilding = $ModeltoRebuildByBuildingNWF | Sort -Unique
        }
    If($NWFList_ByOverall -or $WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
        $ModeltoRebuildByOverallNWF = $NWFList_ByOverall | ForEach-Object { $_.Name }
        $ModeltoRebuildByOverallNWF = $($ModeltoRebuildByOverall;$ModeltoRebuildByOverallNWF)
        $ToIncludeByOverall = $ModeltoRebuildByOverallNWF | Sort -Unique
        }
    If($NWFList_ByFM -or $WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
        $ModeltoRebuildByFMNWF = $NWFList_ByFM | ForEach-Object { $_.Name }
        $ModeltoRebuildByFMNWF = $($ModeltoRebuildByFM;$ModeltoRebuildByFMNWF)
        $ToIncludeByFM = $ModeltoRebuildByFMNWF | Sort -Unique
        }
        
    If($NWFList_ByFM -or $WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
        If($ToIncludeByFM){
            $ModeltoRebuildByFinalNWF = $NWFList_ByFM_All | ForEach-Object { $_.Name }
            $ModeltoRebuildByFinalNWF = $($ModeltoRebuildByFinal;$ModeltoRebuildByFinalNWF)
            $ToIncludeByFinalFM = $ModeltoRebuildByFinalNWF | Sort -Unique
            }
        }
    }

#Writing list of model to be rebuild into text files
$BuildByLevel = $true
$BuildByBuilding = $true
$BuildByOverall = $true
$BuildByFM = $true
$BuildByFinalFM = $true

If($WithUpdatedModel -or $WithNewModel -or $WithRetiredModel){
    try{
        
        #By Level NWF list
        If($ToIncludeByLevel){
            $NWFListByLevel = Get-ChildItem $NWFFolderByLevel -Include $ToIncludeByLevel -Recurse -Filter "*.nwf"
            Write-Output $NWFListByLevel.Name
            WriteLog-Full ("Writing NWF list By Level into: {0}" -f (($ByLevel -split"\\")[-1]))
            Out-File -Filepath $ByLevel -InputObject $NWFListByLevel.FullName
            }
        else{
            WriteLog-Full "All By Level models are up to date."
            Out-File -Filepath $ByLevel -InputObject ""
            $BuildByLevel = $false
            }

        #By Building NWF list
        If($ToIncludeByBuilding){
            $NWFListByBuilding = Get-ChildItem $NWFFolderByBuilding -Include $ToIncludeByBuilding -Recurse -Filter "*.nwf"
            Write-Output $NWFListByBuilding.Name
            WriteLog-Full ("Writing NWF list By Building into: {0}" -f (($ByBuilding -split"\\")[-1]))
            Out-File -Filepath $ByBuilding -InputObject $NWFListByBuilding.FullName
            }
        else{
            WriteLog-Full "All By Building models are up to date."
            Out-File -Filepath $ByBuilding -InputObject ""
            $BuildByBuilding = $false
            }

        #By Overall NWF list
        If($ToIncludeByOverall){
            $NWFListByOverall = Get-ChildItem $NWFFolderByOverall -Include $ToIncludeByOverall -Recurse -Filter "*.nwf"
            Write-Output $NWFListByOverall.Name
            WriteLog-Full ("Writing NWF list By Overall into: {0}" -f (($ByOverall -split"\\")[-1]))
            Out-File -Filepath $ByOverall -InputObject $NWFListByOverall.FullName
            }
        else{
            WriteLog-Full "All By Overall models are up to date."
            Out-File -Filepath $ByOverall -InputObject ""
            $BuildByOverall = $false
            }

        #By FM
        If($ToIncludeByFM){
            $NWFListByFederatedModel = Get-ChildItem $NWFFolderByFM -Include $ToIncludeByFM -Recurse -Filter "*.nwf"
            Write-Output $NWFListByFederatedModel.Name
            WriteLog-Full ("Writing NWF list By Federated Model into: {0}" -f (($ByFederatedModel -split"\\")[-1]))
            Out-File -Filepath $ByFederatedModel -InputObject $NWFListByFederatedModel.FullName
            }
        else{
            WriteLog-Full "All By FM models are up to date."
            Out-File -Filepath $ByFederatedModel -InputObject ""
            $BuildByFM = $false
            }

        #By FM FINAL
        If($ToIncludeByFinalFM){
            $NWFListByFinalFM = Get-ChildItem $NWFFolderByFM -Include $ToIncludeByFinalFM -Recurse -Filter "*.nwf"
            Write-Output $NWFListByFinalFM.Name
            WriteLog-Full ("Writing NWF list By Final FM into: {0}" -f (($ByFinalFM -split"\\")[-1]))
            Out-File -Filepath $ByFinalFM -InputObject $NWFListByFinalFM.FullName
            }
        else{
            WriteLog-Full "All By Final FM models are up to date."
            Out-File -Filepath $ByFinalFM -InputObject ""
            $BuildByFinalFM = $false
            }
        }

    catch{
        $BuildException = $_.Exception.Message
        WriteLog-Full $BuildException -Type ERROR
        $BuildSuccess = $false
     }
}
else{
    WriteLog-Full "Skip rebuild federated models. All models are up to date." -Type INFO
    }

################ NWF BUILD REGION END ################

#Update all nwf searchsets and viewpoints
$NWFList = Get-ChildItem $NWFFolderAll -Exclude "_Archived","_Rejected","test","_Retired","1 By Level" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_Level = Get-ChildItem $NWFFolderByLevel | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
Initialize-NavisworksApi
$napiDC = [Autodesk.Navisworks.Api.Controls.DocumentControl]::new()
WriteLog-Full "Start updating search sets and viewpoints..."
$i = 0

#By Level
try{
    ForEach($nwf in $NWFList_Level){
        $i = $i+1
        Write-Progress -Activity "Cleaning viewpoints for level models..." -Status ("Updating file: {0}" -f $nwf.Name) -PercentComplete (($i/$NWFList_Level.count)*100)
        WriteLog-Full ("Updating file: {0}" -f $nwf)
        $napiDC.Document.TryOpenFile($nwf.FullName)
        $napiDC.Document.SavedViewpoints.Clear()
        $napiDC.Document.SaveFile($nwf.FullName)
        }
 }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
    }

#All other
$i = 0
try{
    if($napiDC.Document.TryOpenFile($SelectionVPfile)) {
        $viewpoint = $napiDC.Document.SavedViewpoints.CreateCopy()
        $selectionset = $napiDC.Document.SelectionSets.CreateCopy()
        ForEach($nwf in $NWFList){
            If(!($nwf.Name -match "(PG-FM|F26-FM)")){
                $i = $i+1
                Write-Progress -Activity "Updating Selection Sets and Viewpoints..." -Status ("Updating file: {0}" -f $nwf.Name) -PercentComplete (($i/$NWFList.count)*100)
                WriteLog-Full ("Updating file: {0}" -f $nwf)
                $napiDC.Document.TryOpenFile($nwf.FullName)
                $napiDC.Document.SavedViewpoints.Clear()
                $napiDC.Document.SavedViewpoints.CopyFrom($viewpoint)
                $napiDC.Document.SelectionSets.Clear()
                $napiDC.Document.SelectionSets.CopyFrom($selectionset)
                $napiDC.Document.SaveFile($nwf.FullName)
            }
        }
        }
    else{
        WriteLog-Full ("Master model with search sets and viewpoints does not exist: {0}" -f $SelectionVPfile) -Type WARN
    }
    $napiDC.Document.Clear()
    $napiDC.Dispose()
 }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
    }

WriteLog-Full "Completed updating search sets and viewpoints..."

################################# FEDERATED MODEL BUILD SECTION #################################

#List all federated models that will be rebuild
WriteLog-Full "Building Federated Model for:"
$ToIncludeByLevel | ForEach-Object {WriteLog-Full "$_"}
$ToIncludeByBuilding | ForEach-Object {WriteLog-Full "$_"}
$ToIncludeByOverall | ForEach-Object {WriteLog-Full "$_"}
$ToIncludeByFM | ForEach-Object {WriteLog-Full "$_"}
$ToIncludeByFinalFM | ForEach-Object {WriteLog-Full "$_"}

#Start building federated model from 5 final text files
try{
    #Building federated model by level
    If($BuildByLevel){
        BuildFederatedModel -Stage Level
        }

    #Building federated model by building
    If($BuildByBuilding){
        BuildFederatedModel -Stage Building
        }

    #Building federated model by overall
    If($BuildByOverall){
        BuildFederatedModel -Stage Overall
        }

    #Building federated model by FEDERATED MODEL
    If($BuildByFM){
        BuildFederatedModel -Stage FM
        }

    #Building federated model by Final FM
    If($BuildByFinalFM){
        BuildFederatedModel -Stage Final
        }
    }
catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
    }

#Copy latest federated model to the NWD folder (ACC folder structure CM DM EM FM)
WriteLog-Full ("Copy latest federated model files to main folder : {0}" -f $MainBuildFolder)

try{
    #By Level Folder
    $ModelType = "CM", "DM", "EM"
    ForEach($Type in $ModelType) {
        $i = 0
        $Files = Get-ChildItem $ByLevelOut -Recurse -Filter ("*-{0}.nwd" -f $Type) | Where-Object { $_.LastWriteTime -gt $DateStarted }
        ForEach($File in $Files){
            $i = $i+1
            $FileDestination = "$MainBuildFolder\{0}\By Level\{1}" -f $Type, $File.Name
            Write-Progress -Activity ("Copying latest Federated Model Files By Level {0}/{1} ({2})..." -f $i, $Files.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
            New-Item -ItemType File -Path $FileDestination -Force
            Copy-Item -Path $File.FullName -Destination $FileDestination -Force
            }
        }

    #By Building Folder
    ForEach($Type in $ModelType) {
        $i = 0
        $Files = Get-ChildItem $ByBuildingOut -Recurse -Filter ("*-{0}.nwd" -f $Type) | Where-Object { $_.LastWriteTime -gt $DateStarted }
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
        $Files = Get-ChildItem $ByOverallOut -Recurse -Filter ("*-{0}.nwd" -f $Type) | Where-Object { $_.LastWriteTime -gt $DateStarted }
        ForEach($File in $Files){
            $i = $i+1
            $FileDestination = "$MainBuildFolder\{0}\By Building\{1}" -f $Type, $File.Name
            Write-Progress -Activity ("Copying latest Federated Model Files By Overall {0}/{1} ({2})..." -f $i, $Files.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
            New-Item -ItemType File -Path $FileDestination -Force
            Copy-Item -Path $File.FullName -Destination $FileDestination -Force
            }
        }

    #By Federated Model Folder
    $i = 0
    $Files = Get-ChildItem $ByFinalFMOut -Recurse -Filter ("*-FM.nwd") | Where-Object { $_.LastWriteTime -gt $DateStarted }
    ForEach($File in $Files){
        $i = $i+1
        $FileDestination = "$MainBuildFolder\FM\{0}" -f $File.Name
        Write-Progress -Activity ("Copying latest Federated Model Files FEDERATED MODEL {0}/{1} ({2})..." -f $i, $Files.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Files.count)*100)
        New-Item -ItemType File -Path $FileDestination -Force
        Copy-Item -Path $File.FullName -Destination $FileDestination -Force
        }
    }
catch{
    $CopyException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
    }

#Send email with the federated model build status and log if there is any error
If(!($BuildSuccess)){
    $DateNow = $((Get-Date).ToString('yyyy-MM-dd'))
    $DateNowFull = Get-Date
    $LogFile = "$LogFolder\Pelican_federated_model_build_log_{0}.csv" -f $DateStartedText

    try{
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.importance = 2
        $mail.subject = "ERROR: Pelican Federated Model Build for $DateNow"
        $mail.body = "There is an error while running the Federated Model Build.`n`nBuild started on <$DateStarted> and finished on <$DateNowFull>"
        $mail.to = "lawrenerno.jinkim@exyte.net;janetjasintha.lopez@exyte.net"
        $mail.Attachments.Add($LogFile)
        WriteLog-Full "Sending email to : lawrenerno.jinkim@exyte.net and janetjasintha.lopez@exyte.net"
        $mail.Send()
        Start-Sleep 20
        $outlook.Quit()
        }

    catch{
        $Exception = $_.Exception.Message
        WriteLog-Full "$Exception" -Type ERROR
        }
    }