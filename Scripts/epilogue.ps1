#load configuration file
. .\Config.ps1

If(!(Test-Path $OverallACCFolder)){
    New-Item -ItemType Directory -Path $OverallACCFolder -Force
    }
If(!(Test-Path $OverallACCFolder_Rejected)){
    New-Item -ItemType Directory -Path $OverallACCFolder_Rejected -Force
    }
If(!(Test-Path $OverallACCFolder_Incorrect)){
    New-Item -ItemType Directory -Path $OverallACCFolder_Incorrect -Force
    }
If(!(Test-Path "$OverallACCFolder\DM")){
    New-Item -ItemType Directory -Path "$OverallACCFolder\DM" -Force
    }
If(!(Test-Path "$OverallACCFolder\CM")){
    New-Item -ItemType Directory -Path "$OverallACCFolder\CM" -Force
    }
If(!(Test-Path "$OverallACCFolder\EM")){
    New-Item -ItemType Directory -Path "$OverallACCFolder\EM" -Force
    }

#Move Rejected models
try{
    #Regex for correct file naming
    $NPattern = "^XPG\w{2,3}-\w{3,4}-\d{3}-.{2}-\w{5}-\w{3}-(CM|DM|EM)(\.|[-_]ROOM\.|[-_]MAXFIT\.|[-_]AMHS\.|-ST\.)nwc$"
    $i = 0
    $Files = Get-ChildItem $OverallACCFolder -Exclude "_Rejected","_Incorrect_folder" | Get-ChildItem -Recurse -Filter "*.nwc"
    If(!($OverallACCFolder_Rejected)){
        New-Item -ItemType Directory -Path $OverallACCFolder_Rejected -Force
        }
    $FList = $Files -notmatch $NPattern
    If($Flist){
        WriteLog-Full ("Moving {0} rejected model into: {1}" -f $Flist.count, $RejectedFolder)
        ForEach($File in $FList){
            $i = $i+1
            Write-Progress -Activity ("Moving rejected models {0}/{1} ({2})..." -f $i, $Flist.count, $File.Name) -Status "Progress: " -PercentComplete (($i/$Flist.count)*100)
            $FFileName = "$OverallACCFolder_Rejected\{0}" -f $File.Name
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

$FilesAll_DM = Get-ChildItem "$OverallACCFolder\DM" -Filter "*.nwc"
$FilesAll_CM = Get-ChildItem "$OverallACCFolder\CM" -Filter "*.nwc"
$FilesAll_EM = Get-ChildItem "$OverallACCFolder\EM" -Filter "*.nwc"
$NFiles_DM = $FilesAll_DM.Name -notlike "*-DM*.nwc"
$NFiles_CM = $FilesAll_CM.Name -notlike "*-CM*.nwc"
$NFiles_EM = $FilesAll_EM.Name -notlike "*-EM*.nwc"

#Check if incorrect folder exist. Create if it does not exist
If(!(Test-Path "$OverallACCFolder_Incorrect")){
    New-Item -ItemType Directory -Path $OverallACCFolder_Incorrect -Force
    }

#Move incorrect file location into folder
If(!($NFiles_DM -eq ('True' -or 'False'))){
    ForEach($f in $NFiles_DM){
        $FFileName = "$OverallACCFolder_Incorrect\{0}" -f $f
        New-Item -ItemType File -Path $FFileName -Force
        Move-Item -Path ("$OverallACCFolder\DM\{0}"-f $f) -Destination $FFileName -Force
        }
    }

Write-Output $NFiles_EM
If(!($NFiles_CM -eq ('True' -or 'False'))){
    ForEach($f in $NFiles_CM){
        $FFileName = "$OverallACCFolder_Incorrect\{0}" -f $f
        New-Item -ItemType File -Path $FFileName -Force
        Move-Item -Path ("$OverallACCFolder\CM\{0}"-f $f) -Destination $FFileName -Force
        }
    }

If(!($NFiles_EM -eq ('True' -or 'False'))){
    ForEach($f in $NFiles_EM){
        $FFileName = "$OverallACCFolder_Incorrect\{0}" -f $f
        New-Item -ItemType File -Path $FFileName -Force
        Move-Item -Path ("$OverallACCFolder\EM\{0}"-f $f) -Destination $FFileName -Force
        }
    }

#Cleanup each folder
If("$OverallACCFolder\DM"){
    Get-ChildItem "$OverallACCFolder\DM" | Where {($_.Extension -ne ".nwc")} | Remove-Item -Force -Recurse
    }
If("$OverallACCFolder\CM"){
    Get-ChildItem "$OverallACCFolder\CM" | Where {($_.Extension -ne ".nwc")} | Remove-Item -Force -Recurse
    }
If("$OverallACCFolder\EM"){
    Get-ChildItem "$OverallACCFolder\EM" | Where {($_.Extension -ne ".nwc")} | Remove-Item -Force -Recurse
    }