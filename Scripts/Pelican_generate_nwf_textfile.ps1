<#
    Program Name : Pelican_generate_nwf_textfiles
    Verison : 1.0.0
    Description : Generate all text files for model builds
#>

. .\Config.ps1

function Show-Menu
{
    param (
        [string]$Title = 'NWCs source folder'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' for Temporary Build Folder."
    Write-Host "2: Press '2' for BIM Machine Build Folder."
    Write-Host "Q: Press 'Q' to quit."
}

$SourceFolder = $TempNWC_All
Show-Menu
$selection = Read-Host "Please make a selection"
switch ($selection)
 {
    '1' {
         $SourceFolder = $TempNWC_All
        } 
    '2' {
         $SourceFolder = $MainNWC_All
        }
    'q' {
         return
     }
 }


Write-Output "Running build on local Computer: $env:computername"
Write-Output "Building all NWF files.."

# ---BY LEVEL---
# ---DM MODEL---

$NWCList = Get-ChildItem $SourceFolder -Exclude "_Archived","_Rejected","test","_Retired","_New.txt" | Get-ChildItem -Recurse -Filter "*.nwc"

$BTextByLevel = "$BatchTextFolder\1 By Level"
$BTextByBuilding = "$BatchTextFolder\2 By Building"
$BTextByOverall = "$BatchTextFolder\3 By Overall"
$BTextByFM = "$BatchTextFolder\FEDERATED MODEL"

New-Item -ItemType Directory -Path "$BTextByLevel" -Force
New-Item -ItemType Directory -Path "$BTextByBuilding" -Force
New-Item -ItemType Directory -Path "$BTextByOverall" -Force
New-Item -ItemType Directory -Path "$BTextByFM" -Force

Write-Output "Building NWF Model (DM) By Level"

#APB1 DM
try{
    $i=0
    ForEach($item in $F26_APB1_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB1_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 DM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB1_DM.$item.FName, $i, $F26_APB1_DM.keys.count) -PercentComplete (($i/$F26_APB1_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#APB2 DM
try{
    $i=0
    ForEach($item in $F26_APB2_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB2_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 DM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB2_DM.$item.FName, $i, $F26_APB2_DM.keys.count) -PercentComplete (($i/$F26_APB2_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#FAB DM
try{
    $i=0
    ForEach($item in $F26_FAB_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_FAB_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_FAB_DM.$item.FName, $i, $F26_FAB_DM.keys.count) -PercentComplete (($i/$F26_FAB_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGB DM
try{
    $i=0
    ForEach($item in $PGB_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGB_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGB_DM.$item.FName, $i, $PGB_DM.keys.count) -PercentComplete (($i/$PGB_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGP DM
try{
    $i=0
    ForEach($item in $PGP_DM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_DM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGP_DM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB DM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGP_DM.$item.FName, $i, $PGP_DM.keys.count) -PercentComplete (($i/$PGP_DM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_DM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }


# ---CM MODEL---

Write-Output "Building NWF Model (CM) By Level"

#APB1 CM
try{
    $i=0
    ForEach($item in $F26_APB1_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB1_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 CM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB1_CM.$item.FName, $i, $F26_APB1_CM.keys.count) -PercentComplete (($i/$F26_APB1_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#APB2 CM
try{
    $i=0
    ForEach($item in $F26_APB2_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB2_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 CM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB2_CM.$item.FName, $i, $F26_APB2_CM.keys.count) -PercentComplete (($i/$F26_APB2_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#FAB CM
try{
    $i=0
    ForEach($item in $F26_FAB_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_FAB_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_FAB_CM.$item.FName, $i, $F26_FAB_CM.keys.count) -PercentComplete (($i/$F26_FAB_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGB CM
try{
    $i=0
    ForEach($item in $PGB_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGB_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGB_CM.$item.FName, $i, $PGB_CM.keys.count) -PercentComplete (($i/$PGB_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGP CM
try{
    $i=0
    ForEach($item in $PGP_CM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_CM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGP_CM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB CM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGP_CM.$item.FName, $i, $PGP_CM.keys.count) -PercentComplete (($i/$PGP_CM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_CM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

# ---EM MODEL---

Write-Output "Building NWF Model (EM) By Level"

#APB1 EM
try{
    $i=0
    ForEach($item in $F26_APB1_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB1_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB1_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB1 EM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB1_EM.$item.FName, $i, $F26_APB1_EM.keys.count) -PercentComplete (($i/$F26_APB1_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB1_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#APB2 EM
try{
    $i=0
    ForEach($item in $F26_APB2_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_APB2_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_APB2_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : APB2 EM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_APB2_EM.$item.FName, $i, $F26_APB2_EM.keys.count) -PercentComplete (($i/$F26_APB2_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_APB2_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#FAB EM
try{
    $i=0
    ForEach($item in $F26_FAB_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $F26_FAB_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $F26_FAB_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $F26_FAB_EM.$item.FName, $i, $F26_FAB_EM.keys.count) -PercentComplete (($i/$F26_FAB_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $F26_FAB_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGB EM
try{
    $i=0
    ForEach($item in $PGB_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGB_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGB_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGB_EM.$item.FName, $i, $PGB_EM.keys.count) -PercentComplete (($i/$PGB_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGB_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGP EM
try{
    $i=0
    ForEach($item in $PGP_EM.keys){
        $i=$i+1
	    $Flist = $NWCList -Match $PGP_EM.$item.FPattern
        If ($Flist){
            Write-Output $Flist.Name
            Write-Output ("Processing file: {0}.txt" -f $PGP_EM.$item.FName)
            Write-Progress -Activity ("Generating NWF By Level : FAB EM") -Status ("Processing file: {0}.txt ({1}\{2})" -f $PGP_EM.$item.FName, $i, $PGP_EM.keys.count) -PercentComplete (($i/$PGP_EM.keys.count)*100)
            $Filein = "$BTextByLevel\{0}.txt" -f $PGP_EM.$item.FName
	        Out-File -Filepath $Filein -InputObject $Flist.FullName
            }
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#---BY BUILDING---
#---DM MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"

Write-Output "Building NWF Model (DM) By Building"

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
    Write-Output ("Processing file: F26_APB1-DM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB1 DM" -Status "Processing file: F26_APB1-DM.txt"
    $Filein = "$BTextByBuilding\F26_APB1-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_APB2-DM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB2 DM" -Status "Processing file: F26_APB2-DM.txt"
    $Filein = "$BTextByBuilding\F26_APB2-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_FAB-DM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: F26_FAB-DM.txt"
    $Filein = "$BTextByBuilding\F26_FAB-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BCS DM
try{
	$Flist = $NWCList -Match $F26_BCS_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $F26_BCS_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LK1 DM
try{
	$Flist = $NWCList -Match $LK1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LK1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGB-DM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: PGB-DM.txt"
    $Filein = "$BTextByBuilding\PGB-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGP-DM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB DM" -Status "Processing file: PGP-DM.txt"
    $Filein = "$BTextByBuilding\PGP-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BG1 DM
try{
	$Flist = $NWCList -Match $BG1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#BG2 DM
try{
	$Flist = $NWCList -Match $BG2_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG2_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LB1 DM
try{
	$Flist = $NWCList -Match $LB1_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LB1_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P09 DM
try{
	$Flist = $NWCList -Match $P09_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P09_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P12 DM
try{
	$Flist = $NWCList -Match $P12_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P12_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGC DM
try{
	$Flist = $NWCList -Match $PGC_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $PGC_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#WTY DM
try{
	$Flist = $NWCList -Match $WTY_DM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $WTY_DM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_DM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#---CM MODEL---

Write-Output "Building NWF Model (CM) By Building"

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
    Write-Output ("Processing file: F26_APB1-CM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB1 CM" -Status "Processing file: F26_APB1-CM.txt"
    $Filein = "$BTextByBuilding\F26_APB1-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_APB2-CM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB2 CM" -Status "Processing file: F26_APB2-CM.txt"
    $Filein = "$BTextByBuilding\F26_APB2-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_FAB-CM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: F26_FAB-CM.txt"
    $Filein = "$BTextByBuilding\F26_FAB-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BCS CM
try{
	$Flist = $NWCList -Match $F26_BCS_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $F26_BCS_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LK1 CM
try{
	$Flist = $NWCList -Match $LK1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LK1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGB-CM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: PGB-CM.txt"
    $Filein = "$BTextByBuilding\PGB-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGP-CM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB CM" -Status "Processing file: PGP-CM.txt"
    $Filein = "$BTextByBuilding\PGP-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BG1 CM
try{
	$Flist = $NWCList -Match $BG1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_CM.FName

        #Add ASU model into the list
        Write-Output ("Adding {0} into the list" -f ($PelicanASUModel -split "\\")[-1])
        $FlistASU = $Flist | ForEach-Object { $_.FullName }
        $FlistASU = $($FlistASU;$PelicanASUModel)
        Out-File -Filepath $Filein -InputObject $FlistASU

        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#BG2 CM
try{
	$Flist = $NWCList -Match $BG2_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG2_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LB1 CM
try{
	$Flist = $NWCList -Match $LB1_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LB1_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P09 CM
try{
	$Flist = $NWCList -Match $P09_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P09_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P12 CM
try{
	$Flist = $NWCList -Match $P12_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P12_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGC CM
try{
	$Flist = $NWCList -Match $PGC_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $PGC_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#WTY CM
try{
	$Flist = $NWCList -Match $WTY_CM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $WTY_CM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_CM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#---EM MODEL---

Write-Output "Building NWF Model (EM) By Building"

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
    Write-Output ("Processing file: F26_APB1-EM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB1 EM" -Status "Processing file: F26_APB1-EM.txt"
    $Filein = "$BTextByBuilding\F26_APB1-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_APB2-EM.txt")
    Write-Progress -Activity "Generating NWF By Building : APB2 EM" -Status "Processing file: F26_APB2-EM.txt"
    $Filein = "$BTextByBuilding\F26_APB2-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: F26_FAB-EM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: F26_FAB-EM.txt"
    $Filein = "$BTextByBuilding\F26_FAB-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BCS EM
try{
	$Flist = $NWCList -Match $F26_BCS_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $F26_BCS_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $F26_BCS_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LK1 EM
try{
	$Flist = $NWCList -Match $LK1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LK1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LK1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGB-EM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: PGB-EM.txt"
    $Filein = "$BTextByBuilding\PGB-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output ("Processing file: PGP-EM.txt")
    Write-Progress -Activity "Generating NWF By Building : FAB EM" -Status "Processing file: PGP-EM.txt"
    $Filein = "$BTextByBuilding\PGP-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BG1 EM
try{
	$Flist = $NWCList -Match $BG1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#BG2 EM
try{
	$Flist = $NWCList -Match $BG2_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $BG2_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $BG2_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#LB1 EM
try{
	$Flist = $NWCList -Match $LB1_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $LB1_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $LB1_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P09 EM
try{
	$Flist = $NWCList -Match $P09_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P09_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P09_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#P12 EM
try{
	$Flist = $NWCList -Match $P12_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $P12_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $P12_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#PGC EM
try{
	$Flist = $NWCList -Match $PGC_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $PGC_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $PGC_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#WTY EM
try{
	$Flist = $NWCList -Match $WTY_EM.FPattern
    If ($Flist){
        Write-Output $Flist.Name
        Write-Output ("Processing file: {0}.txt" -f $WTY_EM.FName)
        $Filein = "$BTextByBuilding\{0}.txt" -f $WTY_EM.FName
	    Out-File -Filepath $Filein -InputObject $Flist.FullName
        }
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
    }

#---BY FEDERATED MODEL---

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$ModelType = "CM", "DM", "EM"

Write-Output "Building NWF Model (FM) By Federated Model"

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
    Write-Output "Processing file: F26_APB1-FM.txt"
    $Filein = "$BTextByFM\F26_APB1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: F26_APB2-FM.txt"
    $Filein = "$BTextByFM\F26_APB2-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: F26_FAB-FM.txt"
    $Filein = "$BTextByFM\F26_FAB-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: F26_BCS-FM.txt"
    $Filein = "$BTextByFM\F26_BCS-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: LK1-FM.txt"
    $Filein = "$BTextByFM\LK1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PGB-FM.txt"
    $Filein = "$BTextByFM\PGB-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PGP-FM.txt"
    $Filein = "$BTextByFM\PGP-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: BG1-FM.txt"
    $Filein = "$BTextByFM\BG1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: BG2-FM.txt"
    $Filein = "$BTextByFM\BG2-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: LB1-FM.txt"
    $Filein = "$BTextByFM\LB1-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: P09-FM.txt"
    $Filein = "$BTextByFM\P09-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: P12-FM.txt"
    $Filein = "$BTextByFM\P12-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PGC-FM.txt"
    $Filein = "$BTextByFM\PGC-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: WTY-FM.txt"
    $Filein = "$BTextByFM\WTY-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#BY OVERALL

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$Buildinglist = "BG1-","BG2-","F26_APB1-","F26_APB2-","F26_FAB-","F26_BCS-","LB1-","LK1-","P09-","P12-","PGB-","PGP-","PGC-","WTY-"

Write-Output "Building NWF Model By Overall"

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
    Write-Output "Processing file: PG-DM.txt"
    $Filein = "$BTextByOverall\PG-DM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PG-CM.txt"
    $Filein = "$BTextByOverall\PG-CM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PG-EM.txt"
    $Filein = "$BTextByOverall\PG-EM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }

#OVERALL FM FOR ALPHA (F26)

$NWDList = Get-ChildItem $TempNWDFolder -Exclude "_Archived","_Rejected","test" | Get-ChildItem -Recurse -Filter "*.nwd"
$F26FMlist = "F26_APB1-FM","F26_APB2-FM","F26_FAB-FM","F26_BCS-FM","LK1-FM"

Write-Output "Building NWF Final Model (F26 and PG)"

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
    Write-Output "Processing file: F26-FM.txt"
    $Filein = "$BTextByFM\F26-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
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
    Write-Output "Processing file: PG-FM.txt"
    $Filein = "$BTextByFM\PG-FM.txt"
    Out-File -Filepath $Filein -InputObject $List
    }

catch{
    $BuildException = $_.Exception.Message
    Write-Output "$BuildException"
 }