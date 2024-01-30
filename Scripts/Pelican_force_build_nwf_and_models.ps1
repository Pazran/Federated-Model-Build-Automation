<#
    Program Name : Pelican_force_build_nwf_and_models
    Verison : 5.2.0
    Description : Force build all federated models (force build all NWF and Federated Models)
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

WriteLog-Full "Building all NWF files.."

# ---BY LEVEL---
# ---DM MODEL---

$NWCList = Get-ChildItem $TempNWC_All -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc"

$BTextByLevel = "$BatchTextFolder\1 By Level"
$BTextByBuilding = "$BatchTextFolder\2 By Building"
$BTextByOverall = "$BatchTextFolder\3 By Overall"
$BTextByFM = "$BatchTextFolder\FEDERATED MODEL"

New-Item -ItemType Directory -Path "$BTextByLevel" -Force
New-Item -ItemType Directory -Path "$BTextByBuilding" -Force
New-Item -ItemType Directory -Path "$BTextByOverall" -Force
New-Item -ItemType Directory -Path "$BTextByFM" -Force

WriteLog-Full "Building NWF Model (DM) By Level"

#APB1 DM
try{
    $i=0
    ForEach($item in $F26_APB1_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_DM.$item.FPattern
        If ($Flist){
            Write-Output ("List of files to be appended:`n {0}`n" -f $Flist.Name)
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB1_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 DM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB1_DM.$item.FName, $i, $F26_APB1_DM.keys.count) -PercentComplete (($i/$F26_APB1_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB1_DM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB1_DM.$item.FName, $NWFFolderByLevel, $F26_APB1_DM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_DM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#APB2 DM
try{
    $i=0
    ForEach($item in $F26_APB2_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB2_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 DM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB2_DM.$item.FName, $i, $F26_APB2_DM.keys.count) -PercentComplete (($i/$F26_APB2_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB2_DM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB2_DM.$item.FName, $NWFFolderByLevel, $F26_APB2_DM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_DM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#FAB DM
try{
    $i=0
    ForEach($item in $F26_FAB_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_FAB_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_FAB_DM.$item.FName, $i, $F26_FAB_DM.keys.count) -PercentComplete (($i/$F26_FAB_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_FAB_DM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_FAB_DM.$item.FName, $NWFFolderByLevel, $F26_FAB_DM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_DM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB DM
try{
    $i=0
    ForEach($item in $PGB_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGB_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGB_DM.$item.FName, $i, $PGB_DM.keys.count) -PercentComplete (($i/$PGB_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGB_DM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGB_DM.$item.FName, $NWFFolderByLevel, $PGB_DM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_DM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGP DM
try{
    $i=0
    ForEach($item in $PGP_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGP_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGP_DM.$item.FName, $i, $PGP_DM.keys.count) -PercentComplete (($i/$PGP_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGP_DM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGP_DM.$item.FName, $NWFFolderByLevel, $PGP_DM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_DM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }


# ---CM MODEL---

WriteLog-Full "Building NWF Model (CM) By Level"

#APB1 CM
try{
    $i=0
    ForEach($item in $F26_APB1_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB1_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 CM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB1_CM.$item.FName, $i, $F26_APB1_CM.keys.count) -PercentComplete (($i/$F26_APB1_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB1_CM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB1_CM.$item.FName, $NWFFolderByLevel, $F26_APB1_CM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_CM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#APB2 CM
try{
    $i=0
    ForEach($item in $F26_APB2_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB2_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 CM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB2_CM.$item.FName, $i, $F26_APB2_CM.keys.count) -PercentComplete (($i/$F26_APB2_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB2_CM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB2_CM.$item.FName, $NWFFolderByLevel, $F26_APB2_CM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_CM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#FAB CM
try{
    $i=0
    ForEach($item in $F26_FAB_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_FAB_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_FAB_CM.$item.FName, $i, $F26_FAB_CM.keys.count) -PercentComplete (($i/$F26_FAB_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_FAB_CM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_FAB_CM.$item.FName, $NWFFolderByLevel, $F26_FAB_CM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_CM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB CM
try{
    $i=0
    ForEach($item in $PGB_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGB_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGB_CM.$item.FName, $i, $PGB_CM.keys.count) -PercentComplete (($i/$PGB_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGB_CM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGB_CM.$item.FName, $NWFFolderByLevel, $PGB_CM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_CM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGP CM
try{
    $i=0
    ForEach($item in $PGP_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGP_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGP_CM.$item.FName, $i, $PGP_CM.keys.count) -PercentComplete (($i/$PGP_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGP_CM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGP_CM.$item.FName, $NWFFolderByLevel, $PGP_CM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_CM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

# ---EM MODEL---

WriteLog-Full "Building NWF Model (EM) By Level"

#APB1 EM
try{
    $i=0
    ForEach($item in $F26_APB1_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB1_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 EM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB1_EM.$item.FName, $i, $F26_APB1_EM.keys.count) -PercentComplete (($i/$F26_APB1_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB1_EM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB1_EM.$item.FName, $NWFFolderByLevel, $F26_APB1_EM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_EM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#APB2 EM
try{
    $i=0
    ForEach($item in $F26_APB2_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_APB2_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 EM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_APB2_EM.$item.FName, $i, $F26_APB2_EM.keys.count) -PercentComplete (($i/$F26_APB2_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_APB2_EM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_APB2_EM.$item.FName, $NWFFolderByLevel, $F26_APB2_EM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_EM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#FAB EM
try{
    $i=0
    ForEach($item in $F26_FAB_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $F26_FAB_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $F26_FAB_EM.$item.FName, $i, $F26_FAB_EM.keys.count) -PercentComplete (($i/$F26_FAB_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_F26_FAB_EM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_FAB_EM.$item.FName, $NWFFolderByLevel, $F26_FAB_EM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_EM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB EM
try{
    $i=0
    ForEach($item in $PGB_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGB_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGB_EM.$item.FName, $i, $PGB_EM.keys.count) -PercentComplete (($i/$PGB_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGB_EM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGB_EM.$item.FName, $NWFFolderByLevel, $PGB_EM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_EM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGP EM
try{
    $i=0
    ForEach($item in $PGP_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            WriteLog-Full ("Processing file: {0}.nwf" -f $PGP_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.nwf ({1}\{2})" -f $PGP_EM.$item.FName, $i, $PGP_EM.keys.count) -PercentComplete (($i/$PGP_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            $Arguments_PGP_EM = '/i "{0}\1 By Level\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGP_EM.$item.FName, $NWFFolderByLevel, $PGP_EM.$item.FName
            Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_EM -Wait -NoNewWindow
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#---BY BUILDING---
#---DM MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"

WriteLog-Full "Building NWF Model (DM) By Building"

#APB1 DM
try{
    $List = @()
    ForEach($item in $F26_APB1_DM.keys){
	    $Flist = $NWDList -Match $F26_APB1_DM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB1-DM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB1 DM" -Status "Processing file: F26_APB1-DM.nwf"
    $Filein = "$BTextByBuilding\F26_APB1-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB1_DM = '/i "{0}" /of "{1}\F26_APB1-DM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#APB2 DM
try{
    $List = @()
    ForEach($item in $F26_APB2_DM.keys){
	    $Flist = $NWDList -Match $F26_APB2_DM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB2-DM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB2 DM" -Status "Processing file: F26_APB2-DM.nwf"
    $Filein = "$BTextByBuilding\F26_APB2-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB2_DM = '/i "{0}" /of "{1}\F26_APB2-DM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#FAB DM
try{
    $List = @()
    ForEach($item in $F26_FAB_DM.keys){
	    $Flist = $NWDList -Match $F26_FAB_DM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_FAB-DM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: F26_FAB-DM.nwf"
    $Filein = "$BTextByBuilding\F26_FAB-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_FAB_DM = '/i "{0}" /of "{1}\F26_FAB-DM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BCS DM
try{
	$Flist = $NWCList -Match $F26_BCS_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $F26_BCS_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_F26_BCS_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_BCS_DM.FName, $NWFFolderByBuilding, $F26_BCS_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_BCS_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LK1 DM
try{
	$Flist = $NWCList -Match $LK1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LK1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LK1_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LK1_DM.FName, $NWFFolderByBuilding, $LK1_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LK1_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB DM
try{
    $List = @()
    ForEach($item in $PGB_DM.keys){
	    $Flist = $NWDList -Match $PGB_DM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGB-DM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: PGB-DM.nwf"
    $Filein = "$BTextByBuilding\PGB-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGB_DM = '/i "{0}" /of "{1}\PGB-DM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGP DM
try{
    $List = @()
    ForEach($item in $PGP_DM.keys){
	    $Flist = $NWDList -Match $PGP_DM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGP-DM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: PGP-DM.nwf"
    $Filein = "$BTextByBuilding\PGP-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGP_DM = '/i "{0}" /of "{1}\PGP-DM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BG1 DM
try{
	$Flist = $NWCList -Match $BG1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_BG1_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG1_DM.FName, $NWFFolderByBuilding, $BG1_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG1_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#BG2 DM
try{
	$Flist = $NWCList -Match $BG2_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG2_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_BG2_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG2_DM.FName, $NWFFolderByBuilding, $BG2_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG2_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LB1 DM
try{
	$Flist = $NWCList -Match $LB1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LB1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LB1_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LB1_DM.FName, $NWFFolderByBuilding, $LB1_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LB1_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P09 DM
try{
	$Flist = $NWCList -Match $P09_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P09_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P09_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P09_DM.FName, $NWFFolderByBuilding, $P09_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P09_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P12 DM
try{
	$Flist = $NWCList -Match $P12_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P12_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P12_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P12_DM.FName, $NWFFolderByBuilding, $P12_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P12_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGC DM
try{
	$Flist = $NWCList -Match $PGC_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $PGC_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_PGC_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGC_DM.FName, $NWFFolderByBuilding, $PGC_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGC_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#WTY DM
try{
	$Flist = $NWCList -Match $WTY_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $WTY_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_WTY_DM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $WTY_DM.FName, $NWFFolderByBuilding, $WTY_DM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_WTY_DM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#---CM MODEL---

WriteLog-Full "Building NWF Model (CM) By Building"

#APB1 CM
try{
    $List = @()
    ForEach($item in $F26_APB1_CM.keys){
	    $Flist = $NWDList -Match $F26_APB1_CM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB1-CM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB1 CM" -Status "Processing file: F26_APB1-CM.nwf"
    $Filein = "$BTextByBuilding\F26_APB1-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB1_CM = '/i "{0}" /of "{1}\F26_APB1-CM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#APB2 CM
try{
    $List = @()
    ForEach($item in $F26_APB2_CM.keys){
	    $Flist = $NWDList -Match $F26_APB2_CM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB2-CM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB2 CM" -Status "Processing file: F26_APB2-CM.nwf"
    $Filein = "$BTextByBuilding\F26_APB2-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB2_CM = '/i "{0}" /of "{1}\F26_APB2-CM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#FAB CM
try{
    $List = @()
    ForEach($item in $F26_FAB_CM.keys){
	    $Flist = $NWDList -Match $F26_FAB_CM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_FAB-CM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: F26_FAB-CM.nwf"
    $Filein = "$BTextByBuilding\F26_FAB-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_FAB_CM = '/i "{0}" /of "{1}\F26_FAB-CM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BCS CM
try{
	$Flist = $NWCList -Match $F26_BCS_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $F26_BCS_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_F26_BCS_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_BCS_CM.FName, $NWFFolderByBuilding, $F26_BCS_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_BCS_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LK1 CM
try{
	$Flist = $NWCList -Match $LK1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LK1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LK1_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LK1_CM.FName, $NWFFolderByBuilding, $LK1_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LK1_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB CM
try{
    $List = @()
    ForEach($item in $PGB_CM.keys){
	    $Flist = $NWDList -Match $PGB_CM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGB-CM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: PGB-CM.nwf"
    $Filein = "$BTextByBuilding\PGB-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGB_CM = '/i "{0}" /of "{1}\PGB-CM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGP CM
try{
    $List = @()
    ForEach($item in $PGP_CM.keys){
	    $Flist = $NWDList -Match $PGP_CM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGP-CM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: PGP-CM.nwf"
    $Filein = "$BTextByBuilding\PGP-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGP_CM = '/i "{0}" /of "{1}\PGP-CM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BG1 CM
try{
	$Flist = $NWCList -Match $BG1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_CM.FName
	    
        #Add ASU model into the list
        WriteLog-Full ("Adding {0} into the list" -f ($PelicanASUModel -split "\\")[-1])
        $FlistASU = $Flist | ForEach-Object { $_.FullName }
        $FlistASU = $($FlistASU;$PelicanASUModel)
        Out-File -Filepath $Filein -InputObject $FlistASU

        $Arguments_BG1_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG1_CM.FName, $NWFFolderByBuilding, $BG1_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG1_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#BG2 CM
try{
	$Flist = $NWCList -Match $BG2_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG2_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_BG2_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG2_CM.FName, $NWFFolderByBuilding, $BG2_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG2_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LB1 CM
try{
	$Flist = $NWCList -Match $LB1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LB1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LB1_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LB1_CM.FName, $NWFFolderByBuilding, $LB1_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LB1_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P09 CM
try{
	$Flist = $NWCList -Match $P09_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P09_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P09_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P09_CM.FName, $NWFFolderByBuilding, $P09_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P09_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P12 CM
try{
	$Flist = $NWCList -Match $P12_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P12_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P12_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P12_CM.FName, $NWFFolderByBuilding, $P12_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P12_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGC CM
try{
	$Flist = $NWCList -Match $PGC_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $PGC_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_PGC_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGC_CM.FName, $NWFFolderByBuilding, $PGC_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGC_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#WTY CM
try{
	$Flist = $NWCList -Match $WTY_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $WTY_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_WTY_CM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $WTY_CM.FName, $NWFFolderByBuilding, $WTY_CM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_WTY_CM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#---EM MODEL---

WriteLog-Full "Building NWF Model (EM) By Building"

#APB1 EM
try{
    $List = @()
    ForEach($item in $F26_APB1_EM.keys){
	    $Flist = $NWDList -Match $F26_APB1_EM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB1-EM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB1 EM" -Status "Processing file: F26_APB1-EM.nwf"
    $Filein = "$BTextByBuilding\F26_APB1-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB1_EM = '/i "{0}" /of "{1}\F26_APB1-EM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#APB2 EM
try{
    $List = @()
    ForEach($item in $F26_APB2_EM.keys){
	    $Flist = $NWDList -Match $F26_APB2_EM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_APB2-EM.nwf")
    Write-Progress -Activity "Generating NWF By Building : APB2 EM" -Status "Processing file: F26_APB2-EM.nwf"
    $Filein = "$BTextByBuilding\F26_APB2-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB2_EM = '/i "{0}" /of "{1}\F26_APB2-EM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#FAB EM
try{
    $List = @()
    ForEach($item in $F26_FAB_EM.keys){
	    $Flist = $NWDList -Match $F26_FAB_EM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: F26_FAB-EM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: F26_FAB-EM.nwf"
    $Filein = "$BTextByBuilding\F26_FAB-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_FAB_EM = '/i "{0}" /of "{1}\F26_FAB-EM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BCS EM
try{
	$Flist = $NWCList -Match $F26_BCS_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $F26_BCS_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_F26_BCS_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $F26_BCS_EM.FName, $NWFFolderByBuilding, $F26_BCS_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_BCS_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LK1 EM
try{
	$Flist = $NWCList -Match $LK1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LK1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LK1_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LK1_EM.FName, $NWFFolderByBuilding, $LK1_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LK1_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGB EM
try{
    $List = @()
    ForEach($item in $PGB_EM.keys){
	    $Flist = $NWDList -Match $PGB_EM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGB-EM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: PGB-EM.nwf"
    $Filein = "$BTextByBuilding\PGB-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGB_EM = '/i "{0}" /of "{1}\PGB-EM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGP EM
try{
    $List = @()
    ForEach($item in $PGP_EM.keys){
	    $Flist = $NWDList -Match $PGP_EM.$item.FName
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full ("Processing file: PGP-EM.nwf")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: PGP-EM.nwf"
    $Filein = "$BTextByBuilding\PGP-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGP_EM = '/i "{0}" /of "{1}\PGP-EM.nwf"' -f $Filein, $NWFFolderByBuilding
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BG1 EM
try{
	$Flist = $NWCList -Match $BG1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_BG1_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG1_EM.FName, $NWFFolderByBuilding, $BG1_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG1_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#BG2 EM
try{
	$Flist = $NWCList -Match $BG2_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $BG2_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_BG2_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $BG2_EM.FName, $NWFFolderByBuilding, $BG2_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG2_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#LB1 EM
try{
	$Flist = $NWCList -Match $LB1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $LB1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_LB1_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $LB1_EM.FName, $NWFFolderByBuilding, $LB1_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LB1_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P09 EM
try{
	$Flist = $NWCList -Match $P09_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P09_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P09_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P09_EM.FName, $NWFFolderByBuilding, $P09_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P09_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#P12 EM
try{
	$Flist = $NWCList -Match $P12_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $P12_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_P12_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $P12_EM.FName, $NWFFolderByBuilding, $P12_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P12_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#PGC EM
try{
	$Flist = $NWCList -Match $PGC_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $PGC_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_PGC_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $PGC_EM.FName, $NWFFolderByBuilding, $PGC_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGC_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#WTY EM
try{
	$Flist = $NWCList -Match $WTY_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        WriteLog-Full ("Processing file: {0}.nwf" -f $WTY_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        $Arguments_WTY_EM = '/i "{0}\2 By Building\{1}.txt" /of "{2}\{3}.nwf"' -f $BatchTextFolder, $WTY_EM.FName, $NWFFolderByBuilding, $WTY_EM.FName
        Start-Process $BatchUtilityProcess -ArgumentList $Arguments_WTY_EM -Wait -NoNewWindow
        }
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
    }

#---BY FEDERATED MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$ModelType = "CM", "DM", "EM"

WriteLog-Full "Building NWF Model (FM) By Federated Model"

#APB1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("F26_APB1-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: F26_APB1-FM.nwf"
    $Filein = "$BTextByFM\F26_APB1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB1_FM = '/i "{0}" /of "{1}\F26_APB1-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB1_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#APB2 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("F26_APB2-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: F26_APB2-FM.nwf"
    $Filein = "$BTextByFM\F26_APB2-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_APB2_FM = '/i "{0}" /of "{1}\F26_APB2-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_APB2_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#FAB FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("F26_FAB-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: F26_FAB-FM.nwf"
    $Filein = "$BTextByFM\F26_FAB-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_FAB_FM = '/i "{0}" /of "{1}\F26_FAB-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FAB_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BCS FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("F26_BCS-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: F26_BCS-FM.nwf"
    $Filein = "$BTextByFM\F26_BCS-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_BCS_FM = '/i "{0}" /of "{1}\F26_BCS-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_BCS_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#LK1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("LK1-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: LK1-FM.nwf"
    $Filein = "$BTextByFM\LK1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_LK1_FM = '/i "{0}" /of "{1}\LK1-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LK1_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGB FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("PGB-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PGB-FM.nwf"
    $Filein = "$BTextByFM\PGB-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGB_FM = '/i "{0}" /of "{1}\PGB-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGB_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGP FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("PGP-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PGP-FM.nwf"
    $Filein = "$BTextByFM\PGP-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGP_FM = '/i "{0}" /of "{1}\PGP-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGP_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BG1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("BG1-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: BG1-FM.nwf"
    $Filein = "$BTextByFM\BG1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_BG1_FM = '/i "{0}" /of "{1}\BG1-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG1_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BG2 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("BG2-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: BG2-FM.nwf"
    $Filein = "$BTextByFM\BG2-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_BG2_FM = '/i "{0}" /of "{1}\BG2-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_BG2_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#LB1 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("LB1-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: LB1-FM.nwf"
    $Filein = "$BTextByFM\LB1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_LB1_FM = '/i "{0}" /of "{1}\LB1-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_LB1_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#P09 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("P09-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: P09-FM.nwf"
    $Filein = "$BTextByFM\P09-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_P09_FM = '/i "{0}" /of "{1}\P09-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P09_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#P12 FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("P12-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: P12-FM.nwf"
    $Filein = "$BTextByFM\P12-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_P12_FM = '/i "{0}" /of "{1}\P12-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_P12_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PGC FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("PGC-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PGC-FM.nwf"
    $Filein = "$BTextByFM\PGC-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PGC_FM = '/i "{0}" /of "{1}\PGC-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PGC_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#WTY FM
try{
    $List = @()
    ForEach($item in $ModelType){
	    $Flist = $NWDList -Match ("WTY-{0}" -f $item)
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: WTY-FM.nwf"
    $Filein = "$BTextByFM\WTY-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_WTY_FM = '/i "{0}" /of "{1}\WTY-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_WTY_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#BY OVERALL

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$Buildinglist = "BG1-","BG2-","F26_APB1-","F26_APB2-","F26_FAB-","F26_BCS-","LB1-","LK1-","P09-","P12-","PGB-","PGP-","PGC-","WTY-"

WriteLog-Full "Building NWF Model By Overall"

#PG DM ----WIP
try{
    $List = @()
    ForEach($building in $Buildinglist){
	    $Flist = $NWDList -Match ("{0}DM" -f $building)
        If ($Flist){
            $List += $Flist.FullName
        }
    }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PG-DM.nwf"
    $Filein = "$BTextByOverall\PG-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PG_DM = '/i "{0}" /of "{1}\PG-DM.nwf"' -f $Filein, $NWFFolderByOverall
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_DM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PG CM
try{
    $List = @()
    ForEach($building in $Buildinglist){
	    $Flist = $NWDList -Match ("{0}CM" -f $building)
        If ($Flist){
            $List += $Flist.FullName
        }
    }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PG-CM.nwf"
    $Filein = "$BTextByOverall\PG-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PG_CM = '/i "{0}" /of "{1}\PG-CM.nwf"' -f $Filein, $NWFFolderByOverall
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_CM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PG EM
try{
    $List = @()
    ForEach($building in $Buildinglist){
	    $Flist = $NWDList -Match ("{0}EM" -f $building)
        If ($Flist){
            $List += $Flist.FullName
        }
    }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PG-EM.nwf"
    $Filein = "$BTextByOverall\PG-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PG_EM = '/i "{0}" /of "{1}\PG-EM.nwf"' -f $Filein, $NWFFolderByOverall
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_EM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#OVERALL FM FOR ALPHA (F26)

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$F26FMlist = "F26_APB1-FM","F26_APB2-FM","F26_FAB-FM","F26_BCS-FM","LK1-FM"

WriteLog-Full "Building NWF Final Model (F26 and PG)"

#F26 FM
try{
    $List = @()
    ForEach($item in $F26FMlist){
	    $Flist = $NWDList -Match $item
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: F26-FM.nwf"
    $Filein = "$BTextByFM\F26-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_F26_FM = '/i "{0}" /of "{1}\F26-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_F26_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#PG FM
try{
    $List = @()
    ForEach($building in $Buildinglist){
        $bpattern = "$building{0}" -f "FM"
	    $Flist = $NWDList -Match $bpattern
        If ($Flist){
            $List += $Flist.FullName
            }
        }
    $List = $List | Sort
    Write-Output $List.Name
    WriteLog-Full "Processing file: PG-FM.nwf"
    $Filein = "$BTextByFM\PG-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    $Arguments_PG_FM = '/i "{0}" /of "{1}\PG-FM.nwf"' -f $Filein, $NWFFolderByFM
    Start-Process $BatchUtilityProcess -ArgumentList $Arguments_PG_FM -Wait -NoNewWindow
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#----Writing final text files for batch utility----

$NWFListByLevel = Get-ChildItem $NWFFolderByLevel -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwf"
WriteLog-Full "Generating final updated text files (By Level, Building, Overall, FM)"

#By Level NWF list
try{
    Write-Output $NWFListByLevel.Name
    WriteLog-Full ("Writing NWF list By Level into: {0}" -f (($ByLevel -split"\\")[-1]))
    Out-File -Filepath $ByLevel -InputObject $NWFListByLevel.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

$NWFListByBuilding = Get-ChildItem $NWFFolderByBuilding -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwf"

#By Building NWF list
try{
    Write-Output $NWFListByBuilding.Name
    WriteLog-Full ("Writing NWF list By Building into: {0}" -f (($ByBuilding -split"\\")[-1]))
    Out-File -Filepath $ByBuilding -InputObject $NWFListByBuilding.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

$NWFListByOverall = Get-ChildItem $NWFFolderByOverall -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwf"

#By Overall NWF list
try{
    Write-Output $NWFListByOverall.Name
    WriteLog-Full ("Writing NWF list By Overall into: {0}" -f (($ByOverall -split"\\")[-1]))
    Out-File -Filepath $ByOverall -InputObject $NWFListByOverall.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

$NWFListByFederatedModel = Get-ChildItem $NWFFolderByFM -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwf"

#By FM
try{
    $lst = $NWFListByFederatedModel -notmatch "(PG-FM|F26-FM)"
    Write-Output $lst.Name
    WriteLog-Full ("Writing NWF list By Federated Model into: {0}" -f (($ByFederatedModel -split"\\")[-1]))
    Out-File -Filepath $ByFederatedModel -InputObject $lst.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

#By FM FINAL
try{
    $lst = $NWFListByFederatedModel -match "(PG-FM|F26-FM)"
    Write-Output $lst.Name
    WriteLog-Full ("Writing NWF list By Final FM into: {0}" -f (($ByFinalFM -split"\\")[-1]))
    Out-File -Filepath $ByFinalFM -InputObject $lst.FullName
    }

catch{
    $BuildException = $_.Exception.Message
    WriteLog-Full "$BuildException" -Type ERROR
    $BuildSuccess=$false
 }

$NWFList = Get-ChildItem $NWFFolderAll -Exclude "_Archived","_Rejected","test","1 By Level" | Get-ChildItem -Recurse -Filter "*.nwf"
$NWFList_Level = Get-ChildItem $NWFFolderByLevel | Get-ChildItem -Recurse -Filter "*.nwf"

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
    $BuildSuccess=$false
    }

#Send email with the federated model build status and log if there is any error
If(!($BuildSuccess)){
    . .\Email.ps1
    }