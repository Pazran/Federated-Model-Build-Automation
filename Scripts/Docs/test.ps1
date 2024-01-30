<#
    Program Name : Pelican_main
    Verison : 4.0.0
    Description : Build all Federated Model
    Author : Lawrenerno Jinkim (lawrenerno.jinkim@exyte.net)
#>

. .\Config.ps1

#Append Model By Level by date
function AppendModelNWF-Dynamic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Hashtable]$ModelArray,
        [Parameter(Mandatory)]
        [System.Object[]]$FileList,
        [Parameter(Mandatory)]
        [System.Object[]]$FileListT,
        [Parameter()]
        [System.Object[]]$FileListR,
        [Parameter(Mandatory)]
        [String]$OutFolder,
        [Parameter(Mandatory)]
        [ValidateSet('Level','Building')]
        [String]$Stage,
        [Parameter()]
        [ValidateSet('Main','Ancillary')]
        [String]$BuildingType
    )

    $SortedModelArray = sortAlphabeticallyRecursive $ModelArray
    Switch($Stage){
        Level {
            $StageData = "1 By Level"
            ForEach($item in ($SortedModelArray.keys).GetEnumerator()){
	            $Flist = $FileList -Match $SortedModelArray.$item.FPattern
                $Flist_today = $FileListT -Match $SortedModelArray.$item.FPattern
                $Flist_retired = $FileListR -Match $SortedModelArray.$item.FPattern
                If (($Flist_today) -or ($Flist_retired)){
                    WriteLog-Full ("Processing file: {0}.nwf" -f $SortedModelArray.$item.FName)
                    $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $SortedModelArray.$item.FName
	                Out-File -Filepath $Filein -InputObject $Flist.FullName
                    $Arguments_SortedModelArray = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $SortedModelArray.$item.FName, $OutFolder, $SortedModelArray.$item.FName
                    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_SortedModelArray -Wait -NoNewWindow
                    }
                else{
                    WriteLog-Full ("Skip rebuild - no new or retired nwc for file:,{0}.nwf" -f $SortedModelArray.$item.FName) -Type INFO
                    }
                }
            }
        Building {
            $StageData = "2 By Building"
	        $Flist = $FileList -Match $SortedModelArray.FPattern
            $Flist_today = $FileListT -Match $SortedModelArray.FPattern
            $Flist_retired = $FileListR -Match $SortedModelArray.FPattern
            If (($Flist_today) -or ($Flist_retired)){
                WriteLog-Full ("Processing file: {0}.nwf" -f $SortedModelArray.FName)
                $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $SortedModelArray.FName
	            Out-File -Filepath $Filein -InputObject $Flist.FullName
                $Arguments_SortedModelArray = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $SortedModelArray.FName, $OutFolder, $SortedModelArray.FName
                Start-Process $BatchUtilityProcess -ArgumentList $Arguments_SortedModelArray -Wait -NoNewWindow
                }
            else{
                WriteLog-Full ("Skip rebuild - no new or retired nwc for file:,{0}.nwf" -f $SortedModelArray.FName) -Type INFO
                }
            }
        }
    }

#Build PG-EM
function RebuildPGEM {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Object[]]$BuildingList,
        [Parameter(Mandatory)]
        [System.Object[]]$FileList,
        [Parameter(Mandatory)]
        [System.Object[]]$FileListT,
        [Parameter(Mandatory)]
        [String]$OutFolder
    )
    $List = @()
        ForEach($building in $Buildinglist){
	        $Flist_today = $FileListT -Match ("{0}EM" -f $building)
            If ($Flist_today){
                ForEach($building in $Buildinglist){
	                $Flist = $NWDList -Match ("{0}EM" -f $building)
                    If ($Flist){
                        $List += $Flist.FullName
                        }
                    }
                $List = $List | Sort
                Write-Output $List
                WriteLog-Full "Processing file: PG-EM.nwf"
                $Filein = "$BatchTextFolder\3 By Overall\PG-EM.txt"
                Out-File -Filepath $Filein -InputObject $List
                $Arguments_PG_EM = '/i "{0}" /of "{1}\PG-EM.nwf"' -f $Filein, $OutFolder
                Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_EM -Wait -NoNewWindow
                Break
                }
            }
    }

#Convert object to hashtable
function ConvertPSObjectToHashtable
{
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process
    {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            $collection = @(
                foreach ($object in $InputObject) { ConvertPSObjectToHashtable $object }
            )

            Write-Output -NoEnumerate $collection
        }
        elseif ($InputObject -is [psobject])
        {
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties)
            {
                $hash[$property.Name] = (ConvertPSObjectToHashtable $property.Value).PSObject.BaseObject
            }

            $hash
        }
        else
        {
            $InputObject
        }
    }
}

#Rebuild NWF file if there is new or retired model
function RebuildNWF-Dynamic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Hashtable]$ModelArray,
        [Parameter(Mandatory)]
        [System.Object[]]$FileList,
        [Parameter(Mandatory)]
        [System.Object[]]$FileListT,
        [Parameter(Mandatory)]
        [String]$OutFolder,
        [Parameter(Mandatory)]
        [ValidateSet('Level','Building')]
        [String]$Stage,
        [Parameter()]
        [ValidateSet('Main','Ancillary')]
        [String]$BuildingType
    )

    $SortedModelArray = sortAlphabeticallyRecursive $ModelArray
    Switch($Stage){
        Level {
            $StageData = "1 By Level"
            ForEach($item in ($SortedModelArray.keys).GetEnumerator()){
	            $Flist = $FileList -Match $SortedModelArray.$item.FPattern
                $Flist_today = $FileListT.Name -Match $SortedModelArray.$item.FPattern
                If ($Flist_today){
                    $Flist = $Flist | Sort
                    Write-Output $Flist
                    WriteLog-Full ("Processing file: {0}.nwf" -f $SortedModelArray.$item.FName)
                    $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $SortedModelArray.$item.FName
	                Out-File -Filepath $Filein -InputObject $Flist.FullName
                    $Arguments_SortedModelArray = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $SortedModelArray.$item.FName, $OutFolder, $SortedModelArray.$item.FName
                    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_SortedModelArray -Wait -NoNewWindow
                    }
                else{
                    WriteLog-Full ("Skip rebuild - no new or retired nwc for file:,{0}.nwf" -f $SortedModelArray.$item.FName) -Type INFO
                    }
                }
            }
        Building {
            $StageData = "2 By Building"
            Switch ($BuildingType){
                Ancillary {
	                $Flist = $FileList -Match $SortedModelArray.FPattern
                    $Flist_today = $FileListT.Name -Match $SortedModelArray.FPattern
                    If ($Flist_today){
                        $Flist = $Flist | Sort
                        Write-Output $Flist
                        WriteLog-Full ("Processing file: {0}.nwf" -f $SortedModelArray.FName)
                        $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $SortedModelArray.FName
	                    Out-File -Filepath $Filein -InputObject $Flist.FullName
                        $Arguments_SortedModelArray = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $SortedModelArray.FName, $OutFolder, $SortedModelArray.FName
                        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_SortedModelArray -Wait -NoNewWindow
                        }
                    else{
                        WriteLog-Full ("Skip rebuild - no new or retired nwc for file:,{0}.nwf" -f $SortedModelArray.FName) -Type INFO
                        }
                    }
                Main {
                    $List = @()
                    ForEach($item in $SortedModelArray.keys){
	                    $Flist_today = $FileListT.Name -Match $SortedModelArray.$item.FName
                        If ($Flist_today){
                            $FileName = $SortedModelArray.$item.FName -replace "-.{3,4}-", "-"
                            ForEach($item in $SortedModelArray.keys){
	                        $Flist = $FileList -Match $SortedModelArray.$item.FName
                            If ($Flist){
                                $List += $Flist.FullName
                                }
                            }
                            $List = $List | Sort
                            Write-Output $List
                            WriteLog-Full ("Processing file: {0}.nwf" -f $FileName)
                            $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $FileName
                            Out-File -Filepath $Filein -InputObject $List
                            $Arguments = '/i "{0}" /of "{1}\{2}.nwf"' -f $Filein, $NWFFolderByBuilding, $FileName
                            Start-Process $BatchUtilityProcess -ArgumentList $Arguments -Wait -NoNewWindow
                            Break
                            $List = @()
                            }
                        }
                   }
            }
        }
    }
}

function GetModelToRebuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Hashtable]$ModelArray,
        [Parameter(Mandatory)]
        [System.Object[]]$FileListT,
        [Parameter(Mandatory)]
        [ValidateSet('Level','Building')]
        [String]$Stage,
        [Parameter()]
        [ValidateSet('Main','Ancillary')]
        [String]$BuildingType
        )

        $RebuildModelArray=@()
        $SortedModelArray = sortAlphabeticallyRecursive $ModelArray
        Switch($Stage){
        Level {
            #$StageData = "1 By Level"
            ForEach($item in ($SortedModelArray.keys).GetEnumerator()){
	            #$Flist = $FileList -Match $SortedModelArray.$item.FPattern
                $Flist_today = $FileListT.Name -Match $SortedModelArray.$item.FPattern
                If ($Flist_today){
                    $RebuildModelArray += $SortedModelArray.$item.FName
                    
                    }
                }
            }
        Building {
            $Flist_today = $FileListT.Name -Match $SortedModelArray.FPattern
            If ($Flist_today){
                $RebuildModelArray += $SortedModelArray.FName
                }
            }
        }
        Return $RebuildModelArray
   }


function RebuildNWFFM {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Array]$ModelArray,
        [Parameter(Mandatory)]
        [System.Object[]]$FileList,
        [Parameter(Mandatory)]
        [String]$OutFolder
    )

    $ModelType = "DM","CM","EM"
    $List = @()
    ForEach($a in $ModelArray){
        ForEach($item in $ModelType){
            $Ptrn = $a -replace "FM","$item"
	        $FlistM = $FileList.Name -Match $Ptrn
            If ($FlistM){
                ForEach($item in $ModelType){
	                $Flist = $FileList -Match ($a -replace "FM","$item")
                    If ($Flist){
                        $List += $Flist.FullName
                        }
                    }
                $List = $List | Sort
                Write-Output $List
                WriteLog-Full ("Processing file: {0}.nwf" -f $a)
                $Filein = "$BatchTextFolder\FEDERATED MODEL\{0}.txt" -f $a
                Out-File -Filepath $Filein -InputObject $List
                $Arguments = '/i "{0}" /of "{1}\{2}.nwf"' -f $Filein, $OutFolder, $a
                Start-Process $BatchUtilityProcess -ArgumentList $Arguments -Wait -NoNewWindow
                $List = @()
                Break
                }
            }
        }
    #Return $List
}

function BuildFederatedModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Level','Building','Overall','FM','Final')]
        [String]$Stage
    )
    Switch($Stage){
        Level{
            $StageDataIN = $ByLevel
            $StageDataOUT = $ByLevelOut
            $StageText = "By Level"
            }
        Building{
            $StageDataIN = $ByBuilding
            $StageDataOUT = $ByBuildingOut
            $StageText = "By Building"
            }
        Overall{
            $StageDataIN = $ByOverall
            $StageDataOUT = $ByOverallOut
            $StageText = "By Overall"
            }
        FM{
            $StageDataIN = $ByFederatedModel
            $StageDataOUT = $ByFederatedModelOut
            $StageText = "By Federated Model"
            }
        Final{
            $StageDataIN = $ByFinalFM
            $StageDataOUT = $ByFinalFMOut
            $StageText = "By Final FM"
            }
        }
    $i=0
    $Fin = Get-Content $StageDataIN
    WriteLog-Full ("Building federated model {0}..." -f $StageText)
    ForEach($nwf in $Fin){
        $i=$i+1
        $nwfname = ((($nwf -split "\\")[-1]).Replace(".nwf",".nwd")).Replace(".nwf",".nwd")
        $TempText = "$StageDataOUT\{0}.txt" -f $nwfname
        Out-File -Filepath $TempText -InputObject $nwf
        $Arguments= '/i "{0}" /od "{1}"' -f $TempText, $StageDataOUT
        WriteLog-Full ("Building model: {0}" -f $nwfname)
        Write-Progress -Activity ("Building federated model {0}" -f $StageText) -Status ("Processing file: {0}" -f $nwfname) -PercentComplete (($i/$Fin.count)*100)
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments -Wait -NoNewWindow
        Remove-Item -Path $TempText -Force
        }
}

$F26FMlist = "F26_APB1-FM","F26_APB2-FM","F26_FAB-FM","F26_BCS-FM","LK1-FM"
$ModelArrayLevel = "F26_APB1","F26_APB2","F26_FAB","PGB","PGP"
$ModelArrayAncillaryBuilding = "F26_BCS","BG1","BG2","LB1","LK1","P09","P12","PGC","WTY"
$ModelPhase = "_DM","_CM","_EM"
$WithRetiredModel = $true
$WithNewModel = $true
$WithUpdatedModel = $true
$NWCList = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc"
$NWCList_today = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc" | Where-Object { $_.LastWriteTime -gt $(((Get-Date).AddDays(-1)).ToString('yyyy-MM-dd')) }

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
}

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

#Check if New model text file exist
If(!(Test-Path $NewModelFile)){
    WriteLog-Full ("New model file does not exist: {0}" -f $NewModelFile) -Type WARN
    $WithNewModel = $false
}
else{
    If(!($ListofNewModel = Get-Content $NewModelFile)){
        $WithNewModel = $false
        }
    $ListofNewModel = Get-Content $NewModelFile
    }

#Rebuild NWF model only if there is new or retired model
try{
    #$ListofNewModel = Get-Content $NewModelFile
    If(!($WithRetiredModel) -and !($WithNewModel)){
        WriteLog-Full ("There is no new model file found in: {0}" -f $NewModelFile) -Type INFO
        WriteLog-Full "Skip rebuild NWF files. All NWF is up to date." -Type INFO
        }
    else{
        WriteLog-Full "New model file found:" -Type INFO
        $ListofNewModel | ForEach-Object {WriteLog-Full "$_"}
        WriteLog-Full "Rebuilding nwf files..."
        $ListOfNewAndRetiredModel = @()
        If($ListofNewModel){
            $ListOfNewAndRetiredModel = $ListofNewModel | ForEach-Object {"$_"}
            }
        If($WithRetiredModel){
            $NWCList_retired = Get-ChildItem "$TempNWC_All\_Retired" -Filter "*.nwc"
            #BEFORE FIX
            #$ListOfNewAndRetiredModel += $NWCList_retired.Name
            $ListOfRetiredModel = $NWCList_retired | ForEach-Object {$_.Name}
            $ListOfNewAndRetiredModel = $($ListofNewModel;$ListOfRetiredModel)

        }
        #$ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired" -Include $ListOfNewAndRetiredModel | Get-ChildItem -Recurse -Filter "*.nwc"
        #$ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_New.txt" -Include $ListOfNewAndRetiredModel -Recurse | Where-Object {$_.FullName -notlike "*\_*"}
        $ListOFNWC = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_New.txt" -Include $ListOfNewAndRetiredModel -Recurse | Where-Object {$_.FullName -notlike "*\_Archived*"}
        
        #Build NWF By Level model for DM CM EM
        $NWFModeltoRebuildByLevel = @()
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayLevel){
                $Array = (Get-Variable $Model$Phase).Value
                RebuildNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $ListOFNWC -OutFolder $NWFFolderByLevel -Stage Level
                $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $ListOFNWC -Stage Level
                $NWFModeltoRebuildByLevel += $ModelList
            }
        }

        #Build NWF for Main Building model for DM CM EM
        $NWFList_Temp = Get-ChildItem $NWFFolderByLevel -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
        $NWFtoRebuildByBuildingMain = @()
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

        #Build NWF for Ancillary By Building model for DM CM EM
        ForEach($Phase in $ModelPhase){
            ForEach($Model in $ModelArrayAncillaryBuilding){
                $Array = (Get-Variable $Model$Phase).Value
                RebuildNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $ListOFNWC -OutFolder $NWFFolderByBuilding -Stage Building -BuildingType Ancillary
                $ModelList = GetModelToRebuild -ModelArray $Array -FileListT $ListOFNWC -Stage Building
                $ToRebuildByBuilding += $ModelList
            }
        }

        $ToRebuildByBuildingMainFM = ($NWFtoRebuildByBuildingMain -replace "[CDE]M","FM") | Sort -Unique
        $ToRebuildByBuildingFM = ($ToRebuildByBuilding -replace "[CDE]M","FM") | Sort -Unique
        $FMModelToBeRebuild = $($ToRebuildByBuildingMainFM;$ToRebuildByBuildingFM)

        #Check if need to rebuild FM NWF files
        If($ListOFNWC.Name -like "*-EM*.nwc"){
            If($ListOFNWC.count -le 1){
                WriteLog-Full ("EM model found: {0}" -f $ListOFNWC.Name) -Type INFO
            }
            else{
                WriteLog-Full ("{0} EM model found:" -f ($ListOFNWC.Name -like "*-EM*.nwc").count) -Type INFO
                $EMModelList = ($ListOFNWC.Name -like "*-EM*.nwc") | ForEach-Object { "$_"}
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
     }
}
else{
    WriteLog-Full "Skip rebuild federated models. All models are up to date." -Type INFO
    }