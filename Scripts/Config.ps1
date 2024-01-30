<#
    Program Name : Configuration file
    Verison : 5.2.0
    Description : Config file for the Pelican Federated Model Build Automation
    Author : Lawrenerno Jinkim (lawrenerno.jinkim@exyte.net)
#>

#Read config files for file and folder locations
$ConfigLocation = ([xml](Get-Content "$PSScriptRoot\Config\Pelican_main_config.xml")).Config

$BatchUtilityProcess = $ConfigLocation.BatchUtility
$LogFolder = "$PSScriptRoot\Logs"
$BackupDirectory = $ConfigLocation.BackupLocation
$RejectedFolder = $ConfigLocation.RejectedLocation
$MainBuildFolder = $ConfigLocation.MainBuildFolder
$TempNWDFolder = $ConfigLocation.TempNWDFolder
$NWFFolderAll = $ConfigLocation.NWFFolderAll
$NWFFolderByLevel = $ConfigLocation.NWFFolderByLevel
$NWFFolderByBuilding = $ConfigLocation.NWFFolderByBuilding
$NWFFolderByOverall = $ConfigLocation.NWFFolderByOverall
$NWFFolderByFM = $ConfigLocation.NWFFolderByFM

$BatchTextFolder = $ConfigLocation.BatchTextFolder
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

$NavisworksAPI_API = $ConfigLocation.NAPI_api
$NavisworksAPI_DC = $ConfigLocation.NAPI_dc
$NavisworksAPI_AUTO = $ConfigLocation.NAPI_auto

$ArgumentsByLevel = '/i "{0}" /od "{1}"' -f $ByLevel, $ByLevelOut
$ArgumentsByBuilding = '/i "{0}" /od "{1}"' -f $ByBuilding, $ByBuildingOut
$ArgumentsByOverall = '/i "{0}" /od "{1}"' -f $ByOverall, $ByOverallOut
$ArgumentsByFederatedModel = '/i "{0}" /od "{1}"' -f $ByFederatedModel, $ByFederatedModelOut
$ArgumentsByFinalFM = '/i "{0}" /od "{1}"' -f $ByFinalFM, $ByFinalFMOut

$DateStarted = Get-Date
$DateStartedText = $((Get-Date).ToString('yyyy-MM-dd'))
$SelectionVPfile = "$PSScriptRoot\Ref\XPGPG-Viewpoints_SearchSet.nwd"
$SpecialFile = "$TempNWDFolder"
$BULogFile = "$PSScriptRoot\Logs\BatchUtility_fm_build_log_$DateStartedText.txt"
$NewModelFile = "C:\test-temp-all\20-Federated-Model\203-NWCs\_New.txt"
$PelicanASUModel = @("C:\test-temp-all\20-Federated-Model\204-NWDs\misc\Pelican ASU Model_Utilities Tie In_real world coordinate 20221205.nwd")
$BuildSuccess = $true

#Janet's folder
$OverallACCFolder = "C:\Users\$Env:UserName\Downloads\_OVERALL_ACC"
$OverallACCFolder_Rejected = "C:\Users\$Env:UserName\Downloads\_OVERALL_ACC\_Rejected"
$OverallACCFolder_Incorrect = "C:\Users\$Env:UserName\Downloads\_OVERALL_ACC\_Incorrect_folder"



################ FUNCTIONS REGION START ################

#Read excel file with read mode or refresh mode
function ReadExcelFile
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [String]$SheetName,
        [Parameter(Mandatory)]
        [ValidateSet('Read','Refresh')]
        [String]$Mode
        )

    $x = New-Object -ComObject Excel.Application
    $x.Visible = $false
    $wb = $x.Workbooks.Open($Path)
    $ws = $wb.Sheets.Item($SheetName)
    Switch($Mode){
        #Read all values from the first column
        Read{
            $startRow = 1
            $xlData = @()
            $count = $ws.Cells.Item(65536,1).End(-4162)
            $xlData = for($startRow=1; $startRow -le $count.row; $startRow++){
                $ws.Cells.Item($startRow, 1).Value()
                }
            Write-Output $xlData
            }

        #Refresh query database and save the file
        Refresh{
            $wb.RefreshAll()
            Write-Output 'Database updated..'
            Start-Sleep -Milliseconds 20
            $wb.Save()
            }
    }
    $wb.Close()
    $x.Quit()
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($x)
}

#Write log file (old log function)
function WriteLog
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$LogString,
        [Parameter()]
        [ValidateSet('ERROR', 'WARN', 'INFO')]
        [string]$Type
    )

    $LogFile = "$LogFolder\Pelican_federated_model_build_log_{0}.csv" -f $DateStartedText
    $DateTime = "{0:dd/MM/yy},{0:HH:mm:ss}" -f (Get-Date)
    if(!($Type)){
        $LogMessage = "$Datetime,$LogString"
        }
    $LogMessage = "$Datetime,$Type : $LogString"
    Add-content $LogFile -value "$LogMessage"
}

#Write log file and output to console
function WriteLog-Full
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$LogString,
        [Parameter()]
        [ValidateSet('ERROR', 'WARN', 'INFO')]
        [string]$Type
    )

    $LogFile = "$LogFolder\Pelican_federated_model_build_log_{0}.csv" -f $DateStartedText
    $DateTime = "{0:dd/MM/yy},{0:HH:mm:ss}" -f (Get-Date)
    if(!($Type)){
        $LogMessage = "$Datetime,$LogString"
        $Output = $LogString
        }
    else{
        $LogMessage = "$Datetime,$Type : $LogString"
        $Output = "$Type : $LogString"
    }
    Add-content $LogFile -value "$LogMessage"
    Write-Output $Output.Replace(',',' ')
}

#Initialize Navisworks API
function Initialize-NavisworksApi {
    Add-Type -Path $NavisworksAPI_API
    Add-Type -Path $NavisworksAPI_DC
    Add-Type -Path $NavisworksAPI_AUTO
    [Autodesk.Navisworks.Api.Controls.ApplicationControl]::Initialize()
}

#Sort hashtable
function sortAlphabeticallyRecursive( [Collections.IDictionary] $hashtable ) {

    $ordered = [ordered] @{}

    $sortedKeys = $hashtable.Keys | Sort-Object

    foreach( $key in $sortedKeys ) {
        
        $value = $hashtable[ $key ]

        # If value is hashtable-like
        if( $value -is [Collections.IDictionary] ) {
            # recurse into nested hashtable
            $ordered[ $key ] = sortAlphabeticallyRecursive $value
        } else {
            # store single value
            $ordered[ $key ] = $value
        }
    }

    $ordered  # Output
}

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
        [String]$Stage
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
                    WriteLog-Full ("Skip rebuild - no new or modified or deleted nwc for file:,{0}.nwf" -f $SortedModelArray.$item.FName) -Type INFO
                    }
                }
            }
        Building {
            $StageData = "2 By Building"
	        $Flist = $FileList -Match $SortedModelArray.FPattern
            $Flist_today = $FileListT -Match $SortedModelArray.FPattern
            If (($Flist_today) -or ($Flist_retired)){
                WriteLog-Full ("Processing file: {0}.nwf" -f $SortedModelArray.FName)
                $Filein = "$BatchTextFolder\{0}\{1}.txt" -f $StageData, $SortedModelArray.FName

                #Check if it is BG1-CM.. If true, append Pelican ASU model
                If($SortedModelArray.FName = "BG1-CM"){
                    WriteLog-Full ("Adding {0} into the list" -f ($PelicanASUModel -split "\\")[-1])
                    $FlistASU = $Flist | ForEach-Object { $_.FullName }
                    $FlistASU = $($FlistASU;$PelicanASUModel)
                    Out-File -Filepath $Filein -InputObject $FlistASU
                    }
                else{
                    Out-File -Filepath $Filein -InputObject $Flist.FullName
                    }
                $Arguments_SortedModelArray = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $SortedModelArray.FName, $OutFolder, $SortedModelArray.FName
                Start-Process $BatchUtilityProcess -ArgumentList $Arguments_SortedModelArray -Wait -NoNewWindow
                }
            else{
                WriteLog-Full ("Skip rebuild - no new or modified or deleted nwc for file:,{0}.nwf" -f $SortedModelArray.FName) -Type INFO
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

#Convert object to hashtable (UNUSED)
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
                        
                        #Check if it is BG1-CM.. If true, append Pelican ASU model
                        If($SortedModelArray.FName = "BG1-CM"){
                            WriteLog-Full ("Adding {0} into the list" -f ($PelicanASUModel -split "\\")[-1])
                            $FlistASU = $Flist | ForEach-Object { $_.FullName }
                            $FlistASU = $($FlistASU;$PelicanASUModel)
                            Out-File -Filepath $Filein -InputObject $FlistASU
                            }
                        else{
                            Out-File -Filepath $Filein -InputObject $Flist.FullName
                            }
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


########################## START NAME PATTERN ############################

# -----------------------
#      BY LEVEL
# -----------------------

# --------- DM MODEL --------- 

#APB1 DM
$F26_APB1_DM = @{
    L1 = @{
        FName = "F26_APB1-1YA-DM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-A0S00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L2 = @{
        FName = "F26_APB1-2YA-DM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-A0S00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L3 = @{
        FName = "F26_APB1-3FA-DM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L4 = @{
        FName = "F26_APB1-4FA-DM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L5 = @{
        FName = "F26_APB1-5RA-DM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-A0S00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    }

#APB2 DM
$F26_APB2_DM = @{
    L1 = @{
        FName = "F26_APB2-1YA-DM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-W0AB0-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L2 = @{
        FName = "F26_APB2-2YA-DM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-W0AB0-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L3 = @{
        FName = "F26_APB2-3FA-DM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L4 = @{
        FName = "F26_APB2-4FA-DM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L5 = @{
        FName = "F26_APB2-5RA-DM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-W0AB0-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    }

#FAB DM
$F26_FAB_DM = @{
    L1 = @{
        FName = "F26_FAB-1YA-DM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L2 = @{
        FName = "F26_FAB-2SA-DM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L3 = @{
        FName = "F26_FAB-3FA-DM"
        FPattern = "^XPGF26-3\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L4 = @{
        FName = "F26_FAB-4IA-DM"
        FPattern = "^XPGF26-4\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L5 = @{
        FName = "F26_FAB-5RA-DM"
        FPattern = "^XPGF26-5\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    }

#PGB DM
$PGB_DM = @{
    L1 = @{
        FName = "PGB-010A-DM"
        FPattern = "^XPGPGB-01\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L2 = @{
        FName = "PGB-020A-DM"
        FPattern = "^XPGPGB-02\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L3 = @{
        FName = "PGB-030A-DM"
        FPattern = "^XPGPGB-03\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L4 = @{
        FName = "PGB-040A-DM"
        FPattern = "^XPGPGB-04\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L5 = @{
        FName = "PGB-050A-DM"
        FPattern = "^XPGPGB-05\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L6 = @{
        FName = "PGB-060A-DM"
        FPattern = "^XPGPGB-06\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L7 = @{
        FName = "PGB-070A-DM"
        FPattern = "^XPGPGB-07\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L8 = @{
        FName = "PGB-080A-DM"
        FPattern = "^XPGPGB-08\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L9 = @{
        FName = "PGB-09RA-DM"
        FPattern = "^XPGPGB-09\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    }

#PGP DM
$PGP_DM = @{
    L0 = @{
        FName = "PGP-0UA-DM"
        FPattern = "^XPGPGP-0\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }

    L1 = @{
        FName = "PGP-010A-DM"
        FPattern = "^XPGPGP-01\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L2 = @{
        FName = "PGP-020A-DM"
        FPattern = "^XPGPGP-02\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L3 = @{
        FName = "PGP-030A-DM"
        FPattern = "^XPGPGP-03\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L4 = @{
        FName = "PGP-040A-DM"
        FPattern = "^XPGPGP-04\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L5 = @{
        FName = "PGP-050A-DM"
        FPattern = "^XPGPGP-05\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L6 = @{
        FName = "PGP-060A-DM"
        FPattern = "^XPGPGP-06\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L7 = @{
        FName = "PGP-070A-DM"
        FPattern = "^XPGPGP-07\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L8 = @{
        FName = "PGP-080A-DM"
        FPattern = "^XPGPGP-08\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L9 = @{
        FName = "PGP-090A-DM"
        FPattern = "^XPGPGP-09\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L10 = @{
        FName = "PGP-100A-DM"
        FPattern = "^XPGPGP-10\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L11 = @{
        FName = "PGP-110A-DM"
        FPattern = "^XPGPGP-11\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L12 = @{
        FName = "PGP-120A-DM"
        FPattern = "^XPGPGP-12\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    L13 = @{
        FName = "PGP-13RA-DM"
        FPattern = "^XPGPGP-13\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-DM|-DM-ROOM)\.nwc$"
        }
    }

# --------- CM MODEL --------- 

#APB1 CM
$F26_APB1_CM = @{
    L1 = @{
        FName = "F26_APB1-1YA-CM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-A0S00-\w{3}-CM\.nwc$"
        }
    L2 = @{
        FName = "F26_APB1-2YA-CM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-A0S00-\w{3}-CM\.nwc$"
        }
    L3 = @{
        FName = "F26_APB1-3FA-CM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}-CM\.nwc$"
        }
    L4 = @{
        FName = "F26_APB1-4FA-CM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}-CM\.nwc$"
        }
    L5 = @{
        FName = "F26_APB1-5RA-CM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-A0S00-\w{3}-CM\.nwc$"
        }
    }

#APB2 CM
$F26_APB2_CM = @{
    L1 = @{
        FName = "F26_APB2-1YA-CM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-W0AB0-\w{3}-CM\.nwc$"
        }
    L2 = @{
        FName = "F26_APB2-2YA-CM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-W0AB0-\w{3}-CM\.nwc$"
        }
    L3 = @{
        FName = "F26_APB2-3FA-CM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}-CM\.nwc$"
        }
    L4 = @{
        FName = "F26_APB2-4FA-CM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}-CM\.nwc$"
        }
    L5 = @{
        FName = "F26_APB2-5RA-CM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-W0AB0-\w{3}-CM\.nwc$"
        }
    }

#FAB CM
$F26_FAB_CM = @{
    L1 = @{
        FName = "F26_FAB-1YA-CM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-CM\.nwc$"
        }
    L2 = @{
        FName = "F26_FAB-2SA-CM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-CM\.nwc$"
        }
    L3 = @{
        FName = "F26_FAB-3FA-CM"
        FPattern = "^XPGF26-3\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-CM\.nwc$"
        }
    L4 = @{
        FName = "F26_FAB-4IA-CM"
        FPattern = "^XPGF26-4\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-CM\.nwc$"
        }
    L5 = @{
        FName = "F26_FAB-5RA-CM"
        FPattern = "^XPGF26-5\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-CM\.nwc$"
        }
    }

#PGB CM
$PGB_CM = @{
    L1 = @{
        FName = "PGB-010A-CM"
        FPattern = "^XPGPGB-01\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L2 = @{
        FName = "PGB-020A-CM"
        FPattern = "^XPGPGB-02\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L3 = @{
        FName = "PGB-030A-CM"
        FPattern = "^XPGPGB-03\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L4 = @{
        FName = "PGB-040A-CM"
        FPattern = "^XPGPGB-04\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L5 = @{
        FName = "PGB-050A-CM"
        FPattern = "^XPGPGB-05\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L6 = @{
        FName = "PGB-060A-CM"
        FPattern = "^XPGPGB-06\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L7 = @{
        FName = "PGB-070A-CM"
        FPattern = "^XPGPGB-07\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L8 = @{
        FName = "PGB-080A-CM"
        FPattern = "^XPGPGB-08\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    L9 = @{
        FName = "PGB-09RA-CM"
        FPattern = "^XPGPGB-09\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-CM\.nwc$"
        }
    }

#PGP CM
$PGP_CM = @{
    L0 = @{
        FName = "PGP-0UA-CM"
        FPattern = "^XPGPGP-0\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }

    L1 = @{
        FName = "PGP-010A-CM"
        FPattern = "^XPGPGP-01\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L2 = @{
        FName = "PGP-020A-CM"
        FPattern = "^XPGPGP-02\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L3 = @{
        FName = "PGP-030A-CM"
        FPattern = "^XPGPGP-03\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L4 = @{
        FName = "PGP-040A-CM"
        FPattern = "^XPGPGP-04\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L5 = @{
        FName = "PGP-050A-CM"
        FPattern = "^XPGPGP-05\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L6 = @{
        FName = "PGP-060A-CM"
        FPattern = "^XPGPGP-06\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L7 = @{
        FName = "PGP-070A-CM"
        FPattern = "^XPGPGP-07\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L8 = @{
        FName = "PGP-080A-CM"
        FPattern = "^XPGPGP-08\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L9 = @{
        FName = "PGP-090A-CM"
        FPattern = "^XPGPGP-09\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L10 = @{
        FName = "PGP-100A-CM"
        FPattern = "^XPGPGP-10\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L11 = @{
        FName = "PGP-110A-CM"
        FPattern = "^XPGPGP-11\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L12 = @{
        FName = "PGP-120A-CM"
        FPattern = "^XPGPGP-12\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    L13 = @{
        FName = "PGP-13RA-CM"
        FPattern = "^XPGPGP-13\w{2}-\d{3}-\w{2}-F0Q00-\w{3}(-CM|-CM-ST)\.nwc$"
        }
    }

# --------- EM MODEL --------- 

#APB1 EM
$F26_APB1_EM = @{
    L1 = @{
        FName = "F26_APB1-1YA-EM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-A0S00-\w{3}-EM\.nwc$"
        }
    L2 = @{
        FName = "F26_APB1-2YA-EM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-A0S00-\w{3}-EM\.nwc$"
        }
    L3 = @{
        FName = "F26_APB1-3FA-EM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}-EM\.nwc$"
        }
    L4 = @{
        FName = "F26_APB1-4FA-EM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-A0S00-\w{3}-EM\.nwc$"
        }
    L5 = @{
        FName = "F26_APB1-5RA-EM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-A0S00-\w{3}-EM\.nwc$"
        }
    }

#APB2 EM
$F26_APB2_EM = @{
    L1 = @{
        FName = "F26_APB2-1YA-EM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-W0AB0-\w{3}-EM\.nwc$"
        }
    L2 = @{
        FName = "F26_APB2-2YA-EM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-W0AB0-\w{3}-EM\.nwc$"
        }
    L3 = @{
        FName = "F26_APB2-3FA-EM"
        FPattern = "^XPGF26-3F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}-EM\.nwc$"
        }
    L4 = @{
        FName = "F26_APB2-4FA-EM"
        FPattern = "^XPGF26-4F[ABJ]-\d{3}-\w{2}-W0AB0-\w{3}-EM\.nwc$"
        }
    L5 = @{
        FName = "F26_APB2-5RA-EM"
        FPattern = "^XPGF26-5[RP]A-\d{3}-\w{2}-W0AB0-\w{3}-EM\.nwc$"
        }
    }

#FAB EM
$F26_FAB_EM = @{
    L1 = @{
        FName = "F26_FAB-1YA-EM"
        FPattern = "^XPGF26-[01]\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-EM\.nwc$"
        }
    L2 = @{
        FName = "F26_FAB-2SA-EM"
        FPattern = "^XPGF26-2\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-EM\.nwc$"
        }
    L3 = @{
        FName = "F26_FAB-3FA-EM"
        FPattern = "^XPGF26-3\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-EM\.nwc$"
        }
    L4 = @{
        FName = "F26_FAB-4IA-EM"
        FPattern = "^XPGF26-4\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-EM\.nwc$"
        }
    L5 = @{
        FName = "F26_FAB-5RA-EM"
        FPattern = "^XPGF26-5\w{2}-\d{3}-\w{2}-[CT]0V00-\w{3}-EM\.nwc$"
        }
    }

#PGB EM
$PGB_EM = @{
    L1 = @{
        FName = "PGB-010A-EM"
        FPattern = "^XPGPGB-01\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L2 = @{
        FName = "PGB-020A-EM"
        FPattern = "^XPGPGB-02\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L3 = @{
        FName = "PGB-030A-EM"
        FPattern = "^XPGPGB-03\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L4 = @{
        FName = "PGB-040A-EM"
        FPattern = "^XPGPGB-04\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L5 = @{
        FName = "PGB-050A-EM"
        FPattern = "^XPGPGB-05\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L6 = @{
        FName = "PGB-060A-EM"
        FPattern = "^XPGPGB-06\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L7 = @{
        FName = "PGB-070A-EM"
        FPattern = "^XPGPGB-07\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L8 = @{
        FName = "PGB-080A-EM"
        FPattern = "^XPGPGB-08\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    L9 = @{
        FName = "PGB-09RA-EM"
        FPattern = "^XPGPGB-09\w{2}-\d{3}-\w{2}-[AB]0[0Q]00-\w{3}-EM\.nwc$"
        }
    }

#PGP EM
$PGP_EM = @{
    L0 = @{
        FName = "PGP-0UA-EM"
        FPattern = "^XPGPGP-0\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }

    L1 = @{
        FName = "PGP-010A-EM"
        FPattern = "^XPGPGP-01\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L2 = @{
        FName = "PGP-020A-EM"
        FPattern = "^XPGPGP-02\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L3 = @{
        FName = "PGP-030A-EM"
        FPattern = "^XPGPGP-03\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L4 = @{
        FName = "PGP-040A-EM"
        FPattern = "^XPGPGP-04\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L5 = @{
        FName = "PGP-050A-EM"
        FPattern = "^XPGPGP-05\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L6 = @{
        FName = "PGP-060A-EM"
        FPattern = "^XPGPGP-06\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L7 = @{
        FName = "PGP-070A-EM"
        FPattern = "^XPGPGP-07\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L8 = @{
        FName = "PGP-080A-EM"
        FPattern = "^XPGPGP-08\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L9 = @{
        FName = "PGP-090A-EM"
        FPattern = "^XPGPGP-09\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L10 = @{
        FName = "PGP-100A-EM"
        FPattern = "^XPGPGP-10\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L11 = @{
        FName = "PGP-110A-EM"
        FPattern = "^XPGPGP-11\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L12 = @{
        FName = "PGP-120A-EM"
        FPattern = "^XPGPGP-12\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    L13 = @{
        FName = "PGP-13RA-EM"
        FPattern = "^XPGPGP-13\w{2}-\d{3}-\w{2}-F0Q00-\w{3}-EM\.nwc$"
        }
    }

# -----------------------
#      BY BUILDING
# -----------------------

#---DM MODEL---

#BG1 DM
$BG1_DM = @{
    FName = "BG1-DM"
    FPattern = "^XPGBG1-[123A][ARM0][AWL]-\d{3}-\w{2}-[AC]0[BD]00-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#BG2 DM
$BG2_DM = @{
    FName = "BG2-DM"
    FPattern = "^XPGBG2-[1A][01A][A]-\d{3}-\w{2}-A0000-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#LB1 DM
$LB1_DM = @{
    FName = "LB1-DM"
    FPattern = "^XPGLB1-([0-4]|A)[ARU0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#P09 DM
$P09_DM = @{
    FName = "P09-DM"
    FPattern = "^XPGP09-([1-4]|A)[ARM0][AEH]-\d{3}-\w{2}-[A-D]0([A-D]|0)00-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#P12 DM
$P12_DM = @{
    FName = "P12-DM"
    FPattern = "^XPGP12-([1-4]|A)[AR0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#PGC DM
$PGC_DM = @{
    FName = "PGC-DM"
    FPattern = "^XPG(PGC|GH[1234])-\w{3}-\d{3}-\w{2}-\w{5}-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#WTY DM
$WTY_DM = @{
    FName = "WTY-DM"
    FPattern = "^XPGWTY-\w{3}-\d{3}-\w{2}-A0000-\w{3}(-DM|-DM-ROOM)\.nwc$"
    }

#BCS DM (this pattern is wrong but right. magic)
$F26_BCS_DM = @{
    FName = "F26_BCS-DM"
    FPattern = "^XPGF26-([1-3]|A)\w{2}-\d{3}-\w{2}-AC0AE-\w{3}-DM\.nwc$"
    }

#LK1 DM
$LK1_DM = @{
    FName = "LK1-DM"
    FPattern = "^XPGLK1-([1-5]|A)\w{2}-\d{3}-\w{2}-A0B00-\w{3}-DM\.nwc$"
    }

#---CM MODEL---

#BG1 CM
$BG1_CM = @{
    FName = "BG1-CM"
    FPattern = "(^XPGBG1-[123A][ARM0][AWL]-\d{3}-\w{2}-[AC]0[BD]00-\w{3}-CM\.nwc$)"
    }

#BG2 CM
$BG2_CM = @{
    FName = "BG2-CM"
    FPattern = "^XPGBG2-[1A][01A][A]-\d{3}-\w{2}-A0000-\w{3}-CM\.nwc$"
    }

#LB1 CM
$LB1_CM = @{
    FName = "LB1-CM"
    FPattern = "^XPGLB1-([0-4]|A)[ARU0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}-CM\.nwc$"
    }

#P09 CM
$P09_CM = @{
    FName = "P09-CM"
    FPattern = "^XPGP09-([1-4]|A)[ARM0][AEH]-\d{3}-\w{2}-[A-D]0([A-D]|0)00-\w{3}-CM\.nwc$"
    }

#P12 CM
$P12_CM = @{
    FName = "P12-CM"
    FPattern = "^XPGP12-([1-4]|A)[AR0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}-CM\.nwc$"
    }

#PGC CM
$PGC_CM = @{
    FName = "PGC-CM"
    FPattern = "^XPG(PGC|GH[1234])-\w{3}-\d{3}-\w{2}-\w{5}-\w{3}-CM\.nwc$"
    }

#WTY CM
$WTY_CM = @{
    FName = "WTY-CM"
    FPattern = "^XPGWTY-\w{3}-\d{3}-\w{2}-A0000-\w{3}-CM\.nwc$"
    }

#BCS CM
$F26_BCS_CM = @{
    FName = "F26_BCS-CM"
    FPattern = "^XPGF26-([1-3]|A)\w{2}-\d{3}-\w{2}-AC0AE-\w{3}-CM\.nwc$"
    }

#LK1 CM
$LK1_CM = @{
    FName = "LK1-CM"
    FPattern = "^XPGLK1-([1-5]|A)\w{2}-\d{3}-\w{2}-A0B00-\w{3}-CM\.nwc$"
    }

#---EM MODEL---

#BG1 EM
$BG1_EM = @{
    FName = "BG1-EM"
    FPattern = "^XPGBG1-[123A][ARM0][AWL]-\d{3}-\w{2}-[AC]0[BD]00-\w{3}-EM\.nwc$"
    }

#BG2 EM
$BG2_EM = @{
    FName = "BG2-EM"
    FPattern = "^XPGBG2-[1A][01A][A]-\d{3}-\w{2}-A0000-\w{3}-EM\.nwc$"
    }

#LB1 EM
$LB1_EM = @{
    FName = "LB1-EM"
    FPattern = "^XPGLB1-([0-4]|A)[ARU0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}-EM\.nwc$"
    }

#P09 EM
$P09_EM = @{
    FName = "P09-EM"
    FPattern = "^XPGP09-([1-4]|A)[ARM0][AEH]-\d{3}-\w{2}-[A-D]0([A-D]|0)00-\w{3}-EM\.nwc$"
    }

#P12 EM
$P12_EM = @{
    FName = "P12-EM"
    FPattern = "^XPGP12-([1-4]|A)[AR0]A-\d{3}-\w{2}-[A-H]0([A-H]|0)00-\w{3}-EM\.nwc$"
    }

#PGC EM
$PGC_EM = @{
    FName = "PGC-EM"
    FPattern = "XPG(PGC|GH[1234])-\w{3}-\d{3}-\w{2}-\w{5}-\w{3}-EM\.nwc$"
    }

#WTY EM
$WTY_EM = @{
    FName = "WTY-EM"
    FPattern = "^XPGWTY-\w{3}-\d{3}-\w{2}-A0000-\w{3}-EM\.nwc$"
    }

#BCS EM
$F26_BCS_EM = @{
    FName = "F26_BCS-EM"
    FPattern = "^XPGF26-([1-3]|A)\w{2}-\d{3}-\w{2}-AC0AE-\w{3}-EM\.nwc$"
    }

#LK1 EM
$LK1_EM = @{
    FName = "LK1-EM"
    FPattern = "^XPGLK1-([1-5]|A)\w{2}-\d{3}-\w{2}-A0B00-\w{3}-EM\.nwc$"
    }

########################## END NAME PATTERN ############################