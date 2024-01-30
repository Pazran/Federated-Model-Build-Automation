<#
    Program Name : Pelican_main_conservative
    Verison : 5.2.0
    Description : Build modified, new and retired (build all affected NWF and Federated Model)
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
    $BuildSuccess = $false
    }

#Copy the latest NWCs into the temporary build folder
try{
    WriteLog-Full "Copy UPDATED and NEW NWCs into temporary build folder"
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
}
catch{
    $CopyNWCException = $_.Exception.Message
    WriteLog-Full "$CopyNWCException" -Type ERROR
    $BuildSuccess = $false
}

#Move Rejected models
try{
    #Regex for correct file naming
    $NPattern = "^XPG\w{2,3}-\w{3,4}-\d{3}-.{2}-\w{5}-\w{3}-(CM|DM|EM)(\.|[-_]ROOM\.|[-_]MAXFIT\.|[-_]AMHS\.|-ST\.)nwc$"
    $i = 0
    $Files = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","_Retired" | Get-ChildItem  -Recurse -FIlter "*.nwc"
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
    $BuildSuccess = $false
    }

# ---BY LEVEL---
# ---DM MODEL---

$BTextByLevel = "$BatchTextFolder\1 By Level"
$BTextByBuilding = "$BatchTextFolder\2 By Building"
$BTextByOverall = "$BatchTextFolder\3 By Overall"
$BTextByFM = "$BatchTextFolder\FEDERATED MODEL"

New-Item -ItemType Directory -Path "$BTextByLevel" -Force
New-Item -ItemType Directory -Path "$BTextByBuilding" -Force
New-Item -ItemType Directory -Path "$BTextByOverall" -Force
New-Item -ItemType Directory -Path "$BTextByFM" -Force

WriteLog-Full "Building NWF Model By Level.."

$ModelArrayLevel = "F26_APB1","F26_APB2","F26_FAB","PGB","PGP"
$ModelArrayLevelBuilding = "F26_BCS","BG1","BG2","LB1","LK1","P09","P12","PGC","WTY"
$ModelPhase = "_DM","_CM","_EM"
$WithRetiredModel = $true
$NWCList = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwc"
$NWCList_today = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwc" | Where-Object { $_.LastWriteTime -gt $(((Get-Date).AddDays(-1)).ToString('yyyy-MM-dd')) }

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
                WriteLog-Full ("Retired model: {}" -f $r.Name)
                }
            }
    }

#Build By Level model for DM CM EM
try{
    ForEach($Phase in $ModelPhase){
        ForEach($Model in $ModelArrayLevel){
            $Array = (Get-Variable $Model$Phase).Value
            If($WithRetiredModel){
                AppendModelNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $NWCList_today -FileListR $NWCList_retired -OutFolder $NWFFolderByLevel -Stage Level
                }
            else{
                AppendModelNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $NWCList_today -OutFolder $NWFFolderByLevel -Stage Level
                }
        }
    }
}

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess = $false
    }

#---BY BUILDING---
#---DM MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwd"
$NWFList_today = Get-ChildItem $NWFFolderAll -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf"| Where-Object { $_.LastWriteTime -gt $DateStarted }

WriteLog-Full "Building NWF Model (DM) By Building.."

#APB1 DM
try{
    $List = @()
    ForEach($item in $F26_APB1_DM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB1_DM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB1_DM.keys){
	        $Flist = $NWDList -Match $F26_APB1_DM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB1-DM.nwf"
            $Filein = "$BTextByBuilding\F26_APB1-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB1_DM = '/i "{0}" /of "{1}\F26_APB1-DM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#APB2 DM
try{
    $List = @()
    ForEach($item in $F26_APB2_DM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB2_DM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB2_DM.keys){
	        $Flist = $NWDList -Match $F26_APB2_DM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB2-DM.nwf"
            $Filein = "$BTextByBuilding\F26_APB2-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB2_DM = '/i "{0}" /of "{1}\F26_APB2-DM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#FAB DM
try{
    $List = @()
    ForEach($item in $F26_FAB_DM.keys){
	    $Flist_today = $NWFList_today -Match $F26_FAB_DM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_FAB_DM.keys){
	        $Flist = $NWDList -Match $F26_FAB_DM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_FAB-DM.nwf"
            $Filein = "$BTextByBuilding\F26_FAB-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_FAB_DM = '/i "{0}" /of "{1}\F26_FAB-DM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGB DM
try{
    $List = @()
    ForEach($item in $PGB_DM.keys){
	    $Flist_today = $NWFList_today -Match $PGB_DM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGB_DM.keys){
	        $Flist = $NWDList -Match $PGB_DM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGB-DM.nwf"
            $Filein = "$BTextByBuilding\PGB-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGB_DM = '/i "{0}" /of "{1}\PGB-DM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGP DM
try{
    $List = @()
    ForEach($item in $PGP_DM.keys){
	    $Flist_today = $NWFList_today -Match $PGP_DM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGP_DM.keys){
	        $Flist = $NWDList -Match $PGP_DM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGP-DM.nwf"
            $Filein = "$BTextByBuilding\PGP-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGP_DM = '/i "{0}" /of "{1}\PGP-DM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#---CM MODEL---

WriteLog-Full "Building NWF Model (CM) By Building"

#APB1 CM
try{
    $List = @()
    ForEach($item in $F26_APB1_CM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB1_CM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB1_CM.keys){
	        $Flist = $NWDList -Match $F26_APB1_CM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB1-CM.nwf"
            $Filein = "$BTextByBuilding\F26_APB1-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB1_CM = '/i "{0}" /of "{1}\F26_APB1-CM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#APB2 CM
try{
    $List = @()
    ForEach($item in $F26_APB2_CM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB2_CM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB2_CM.keys){
	        $Flist = $NWDList -Match $F26_APB2_CM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB2-CM.nwf"
            $Filein = "$BTextByBuilding\F26_APB2-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB2_CM = '/i "{0}" /of "{1}\F26_APB2-CM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#FAB CM
try{
    $List = @()
    ForEach($item in $F26_FAB_CM.keys){
	    $Flist_today = $NWFList_today -Match $F26_FAB_CM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_FAB_CM.keys){
	        $Flist = $NWDList -Match $F26_FAB_CM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_FAB-CM.nwf"
            $Filein = "$BTextByBuilding\F26_FAB-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_FAB_CM = '/i "{0}" /of "{1}\F26_FAB-CM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGB CM
try{
    $List = @()
    ForEach($item in $PGB_CM.keys){
	    $Flist_today = $NWFList_today -Match $PGB_CM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGB_CM.keys){
	        $Flist = $NWDList -Match $PGB_CM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGB-CM.nwf"
            $Filein = "$BTextByBuilding\PGB-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGB_CM = '/i "{0}" /of "{1}\PGB-CM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGP CM
try{
    $List = @()
    ForEach($item in $PGP_CM.keys){
	    $Flist_today = $NWFList_today -Match $PGP_CM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGP_CM.keys){
	        $Flist = $NWDList -Match $PGP_CM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGP-CM.nwf"
            $Filein = "$BTextByBuilding\PGP-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGP_CM = '/i "{0}" /of "{1}\PGP-CM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#---EM MODEL---

WriteLog-Full "Building NWF Model (EM) By Building"

#APB1 EM
try{
    $List = @()
    ForEach($item in $F26_APB1_EM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB1_EM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB1_EM.keys){
	        $Flist = $NWDList -Match $F26_APB1_EM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB1-EM.nwf"
            $Filein = "$BTextByBuilding\F26_APB1-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB1_EM = '/i "{0}" /of "{1}\F26_APB1-EM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#APB2 EM
try{
    $List = @()
    ForEach($item in $F26_APB2_EM.keys){
	    $Flist_today = $NWFList_today -Match $F26_APB2_EM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_APB2_EM.keys){
	        $Flist = $NWDList -Match $F26_APB2_EM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB2-EM.nwf"
            $Filein = "$BTextByBuilding\F26_APB2-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB2_EM = '/i "{0}" /of "{1}\F26_APB2-EM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#FAB EM
try{
    $List = @()
    ForEach($item in $F26_FAB_EM.keys){
	    $Flist_today = $NWFList_today -Match $F26_FAB_EM.$item.FName
        If ($Flist_today){
            ForEach($item in $F26_FAB_EM.keys){
	        $Flist = $NWDList -Match $F26_FAB_EM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_FAB-EM.nwf"
            $Filein = "$BTextByBuilding\F26_FAB-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_FAB_EM = '/i "{0}" /of "{1}\F26_FAB-EM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGB EM
try{
    $List = @()
    ForEach($item in $PGB_EM.keys){
	    $Flist_today = $NWFList_today -Match $PGB_EM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGB_EM.keys){
	        $Flist = $NWDList -Match $PGB_EM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGB-EM.nwf"
            $Filein = "$BTextByBuilding\PGB-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGB_EM = '/i "{0}" /of "{1}\PGB-EM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGP EM
try{
    $List = @()
    ForEach($item in $PGP_EM.keys){
	    $Flist_today = $NWFList_today -Match $PGP_EM.$item.FName
        If ($Flist_today){
            ForEach($item in $PGP_EM.keys){
	        $Flist = $NWDList -Match $PGP_EM.$item.FName
            If ($Flist){
                $List += $Flist.FullName
                }
            }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGP-EM.nwf"
            $Filein = "$BTextByBuilding\PGP-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGP_EM = '/i "{0}" /of "{1}\PGP-EM.nwf"' -f $Filein, $NWFFolderByBuilding
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }


#Build Ancillary By Building model for DM CM EM

WriteLog-Full "Building NWF Model - Ancillary Buildings"
try{
    ForEach($Phase in $ModelPhase){
        ForEach($Model in $ModelArrayLevelBuilding){
            $Array = (Get-Variable $Model$Phase).Value
            If($WithRetiredModel){
                AppendModelNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $NWCList_today -FileListR $NWCList_retired -OutFolder $NWFFolderByBuilding -Stage Building
                }
            else{
                AppendModelNWF-Dynamic -ModelArray $Array -FileList $NWCList -FileListT $NWCList_today -OutFolder $NWFFolderByBuilding -Stage Building
                }
        }
    }
}

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess = $false
    }

#---BY FEDERATED MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwd"
$NWFList_today = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf"| Where-Object { $_.LastWriteTime -gt $DateStarted }
$ModelType = "CM", "DM", "EM"

WriteLog-Full "Building NWF Model By Federated Model"

#APB1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("F26_APB1-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("F26_APB1-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB1-FM.nwf"
            $Filein = "$BTextByFM\F26_APB1-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB1_FM = '/i "{0}" /of "{1}\F26_APB1-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#APB2 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("F26_APB2-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("F26_APB2-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_APB2-FM.nwf"
            $Filein = "$BTextByFM\F26_APB2-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_APB2_FM = '/i "{0}" /of "{1}\F26_APB2-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#FAB FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("F26_FAB-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("F26_FAB-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_FAB-FM.nwf"
            $Filein = "$BTextByFM\F26_FAB-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_FAB_FM = '/i "{0}" /of "{1}\F26_FAB-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#BCS FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("F26_BCS-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("F26_BCS-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: F26_BCS-FM.nwf"
            $Filein = "$BTextByFM\F26_BCS-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_F26_BCS_FM = '/i "{0}" /of "{1}\F26_BCS-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_BCS_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#LK1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("LK1-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("LK1-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: LK1-FM.nwf"
            $Filein = "$BTextByFM\LK1-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_LK1_FM = '/i "{0}" /of "{1}\LK1-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LK1_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGB FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("PGB-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("PGB-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGB-FM.nwf"
            $Filein = "$BTextByFM\PGB-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGB_FM = '/i "{0}" /of "{1}\PGB-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGP FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("PGP-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("PGP-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGP-FM.nwf"
            $Filein = "$BTextByFM\PGP-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGP_FM = '/i "{0}" /of "{1}\PGP-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#BG1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("BG1-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("BG1-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: BG1-FM.nwf"
            $Filein = "$BTextByFM\BG1-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_BG1_FM = '/i "{0}" /of "{1}\BG1-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG1_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#BG2 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("BG2-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("BG2-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: BG2-FM.nwf"
            $Filein = "$BTextByFM\BG2-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_BG2_FM = '/i "{0}" /of "{1}\BG2-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG2_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#LB1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("LB1-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("LB1-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: LB1-FM.nwf"
            $Filein = "$BTextByFM\LB1-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_LB1_FM = '/i "{0}" /of "{1}\LB1-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LB1_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#P09 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("P09-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("P09-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: P09-FM.nwf"
            $Filein = "$BTextByFM\P09-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_P09_FM = '/i "{0}" /of "{1}\P09-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P09_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#P12 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("P12-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("P12-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: P12-FM.nwf"
            $Filein = "$BTextByFM\P12-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_P12_FM = '/i "{0}" /of "{1}\P12-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P12_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PGC FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("PGC-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("PGC-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PGC-FM.nwf"
            $Filein = "$BTextByFM\PGC-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PGC_FM = '/i "{0}" /of "{1}\PGC-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGC_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#WTY FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist_today = $NWFList_today -Match ("WTY-{0}" -f $item)
        If ($Flist_today){
            ForEach($item in $ModelType){
	            $Flist = $NWDList -Match ("WTY-{0}" -f $item)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: WTY-FM.nwf"
            $Filein = "$BTextByFM\WTY-FM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_WTY_FM = '/i "{0}" /of "{1}\WTY-FM.nwf"' -f $Filein, $NWFFolderByFM
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_WTY_FM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#BY OVERALL

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwd"
$NWFList_today = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$Buildinglist = "BG1-","BG2-","F26_APB1-","F26_APB2-","F26_FAB-","F26_BCS-","LB1-","LK1-","P09-","P12-","PGB-","PGP-","PGC-","WTY-"

WriteLog-Full "Building NWF Model By Overall"

#PG DM
try{
    $List = @()
    
    ForEach($building in $Buildinglist){
	    $Flist_today = $NWFList_today -Match ("{0}DM" -f $building)
        If ($Flist_today){
            ForEach($building in $Buildinglist){
	            $Flist = $NWDList -Match ("{0}DM" -f $building)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PG-DM.nwf"
            $Filein = "$BTextByOverall\PG-DM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PG_DM = '/i "{0}" /of "{1}\PG-DM.nwf"' -f $Filein, $NWFFolderByOverall
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_DM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PG CM
try{
    $List = @()
    
    ForEach($building in $Buildinglist){
	    $Flist_today = $NWFList_today -Match ("{0}CM" -f $building)
        If ($Flist_today){
            ForEach($building in $Buildinglist){
	            $Flist = $NWDList -Match ("{0}CM" -f $building)
                If ($Flist){
                    $List += $Flist.FullName
                    }
                }
            $List = $List | Sort
            Write-Output $List
            WriteLog-Full "Processing file: PG-CM.nwf"
            $Filein = "$BTextByOverall\PG-CM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PG_CM = '/i "{0}" /of "{1}\PG-CM.nwf"' -f $Filein, $NWFFolderByOverall
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_CM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#PG EM
try{
    $List = @()
    
    ForEach($building in $Buildinglist){
	    $Flist_today = $NWFList_today -Match ("{0}EM" -f $building)
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
            $Filein = "$BTextByOverall\PG-EM.txt"
            Out-File -Filepath $Filein -InputObject $List
            $Arguments_PG_EM = '/i "{0}" /of "{1}\PG-EM.nwf"' -f $Filein, $NWFFolderByOverall
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_EM -Wait -NoNewWindow
            Break
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#----Writing final text files for batch utility----

$NWFListByLevel = Get-ChildItem $NWFFolderByLevel -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }

#By Level NWF list
try{
    Write-Output $NWFListByLevel.Name
    WriteLog-Full ("Writing NWF list By Building into: {0}" -f $ByLevel)
    Out-File -Filepath $ByLevel -InputObject $NWFListByLevel.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

$NWFListByBuilding = Get-ChildItem $NWFFolderByBuilding -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }

#By Building NWF list
try{
    Write-Output $NWFListByBuilding.Name
    WriteLog-Full ("Writing NWF list By Building into: {0}" -f $ByBuilding)
    Out-File -Filepath $ByBuilding -InputObject $NWFListByBuilding.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

$NWFListByOverall = Get-ChildItem $NWFFolderByOverall -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }

#By Overall NWF list
try{
    Write-Output $NWFListByOverall.Name
    WriteLog-Full ("Writing NWF list By Overall into: {0}" -f $ByOverall)
    Out-File -Filepath $ByOverall -InputObject $NWFListByOverall.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

$NWFListByFederatedModel = Get-ChildItem $NWFFolderByFM -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFListByFederatedModelAll = Get-ChildItem $NWFFolderByFM -Exclude "_Archived","_Rejected","test","_Retired" | Get-ChildItem -Recurse -Filter "*.nwf"

#By FM
try{
    $lst = $NWFListByFederatedModel -notmatch "(PG-FM|F26-FM)"
    Write-Output $lst.Name
    WriteLog-Full ("Writing NWF list By Federated Model into: {0}" -f $ByFederatedModel)
    Out-File -Filepath $ByFederatedModel -InputObject $lst.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

#By FM FINAL (Always included in the final text files)
try{
    $lst = $NWFListByFederatedModelAll -match "(PG-FM|F26-FM)"
    Write-Output $lst.Name
    WriteLog-Full ("Writing NWF list By Final FM into: {0}" -f $ByFinalFM)
    Out-File -Filepath $ByFinalFM -InputObject $lst.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
 }

$NWFList = Get-ChildItem $NWFFolderAll -Exclude "_Archived","_Rejected","test","_Retired","1 By Level" | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }
$NWFList_Level = Get-ChildItem $NWFFolderByLevel | Get-ChildItem -Recurse -Filter "*.nwf" | Where-Object { $_.LastWriteTime -gt $DateStarted }

#Update all nwf searchsets and viewpoints
Initialize-NavisworksApi
$napiDC = [Autodesk.Navisworks.Api.Controls.DocumentControl]::new()
$i = 0
WriteLog-Full "Start updating search sets and viewpoints..."

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
try{
    if($napiDC.Document.TryOpenFile($SelectionVPfile)) {
        ForEach($nwf in $NWFList){
            If(!($nwf.Name -match "(PG-FM|F26-FM)")){
                $i = $i+1
                Write-Progress -Activity "Updating Selection Sets and Viewpoints..." -Status ("Updating file: {0}" -f $nwf.Name) -PercentComplete (($i/$NWFList.count)*100)
                WriteLog-Full ("Updating file: {0}" -f $nwf)
                $viewpoint = $napiDC.Document.SavedViewpoints.CreateCopy()
                $selectionset = $napiDC.Document.SelectionSets.CreateCopy()
                $napiDC.Document.TryOpenFile($nwf.FullName)
                $napiDC.Document.SavedViewpoints.Clear()
                $napiDC.Document.SavedViewpoints.CopyFrom($viewpoint)
                $napiDC.Document.SelectionSets.Clear()
                $napiDC.Document.SelectionSets.CopyFrom($selectionset)
                $napiDC.Document.SaveFile($nwf.FullName)
            }
        }
    else{
        WriteLog-Full ("Master model with search sets and viewpoints does not exist: {0}" -f $SelectionVPfile) -Type WARN
        #Write-Host "" -BackgroundColor Red -ForegroundColor Black
    }
    }
 }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess = $false
    }
WriteLog-Full "Completed updating search sets and viewpoints..."

################################# FEDERATED MODEL BUILD SECTION #################################

WriteLog-Full "Building Federated Model..."

#Start building federated model from 5 final text files
try{
    #Building federated model by level
    BuildFederatedModel -Stage Level

    #Building federated model by building
    BuildFederatedModel -Stage Building

    #Building federated model by overall
    BuildFederatedModel -Stage Overall

    #Building federated model by FEDERATED MODEL
    BuildFederatedModel -Stage FM

    #Building federated model by Final FM
    BuildFederatedModel -Stage Final
    }
catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full $BuildException -Type ERROR
    $BuildSuccess=$false
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
    . .\Email.ps1
    }