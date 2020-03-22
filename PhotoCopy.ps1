#Set encoding for error output
& cmd /c ver | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("cp866")

#Clear console
cls

#Ask if camera is connected locally (usb) - skips credentials and connection.
$isUSBconnectedPrompt = Read-Host -Prompt 'Type any key if the camera is connected to USB locally, ENTER otherwise'
$isUSBconnected = $true
if (!$isUSBconnectedPrompt) {$isUSBconnected = $false}

if (!$isUSBconnected) {
    #Get credentials for admin
    $cred = get-Credential -credential 'TRCONT.RU\adm_vlasenkovda'
    $networkCred = $cred.GetNetworkCredential()
    $fullUsername = $networkCred.Domain + '\' + $networkCred.UserName
}

#Choose source
$sourcefolder = Read-Host -Prompt 'Full path to photos'
#$sourcefolder = '\\172.26.8.2\C$\test'

#Choose period
$months = ("Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь")
$askstring = 'Sync for how many days? Press Enter for 31 day (1 month)'
$chosenperiod = Read-Host -Prompt $askstring
if (!$chosenperiod) {$chosenperiod = 31}
$periodfrom = (Get-Date).AddDays(-$chosenperiod).ToString('dd.MM.yyyy')
$periodto = Get-Date -Format 'dd.MM.yyyy'
$printperiod = 'Copying photos from ' + $periodfrom + ' to ' + $periodto
Write-Host $printperiod

#Set destination folder
$destfolder = '\\B92-FILE-03\K\АТО\'
if (!(test-path $destfolder)){
    New-Item -ItemType Directory -Force -Path $destfolder | Out-Null
}

if (!$isUSBconnected){
    #Connect to source
    if([System.IO.File]::Exists($sourcefolder)){
        net use /d $sourcefolder | Out-Null
    }
    Write-Host "Connecting to" $sourcefolder "..."
    net use $sourcefolder /USER:$fullUsername $networkCred.Password
}

#Select only photo and video
$extensions =  '.mp4','.jpg','.jpeg','.png','.bmp','.mov'
$extensions += '.MP4','.JPG','.JPEG','.PNG','.BMP','.MOV'

#Find files in selected period
$Photos = Get-ChildItem -Path $sourcefolder -File -Recurse | Where-Object {($_.LastWriteTime -gt (Get-Date).AddDays(-$chosenperiod)) -and ($extensions.Contains($_.extension))}

#Progress bar stuff
$i = 0
$reallycopied = 0
$speedmeasurepoint = 10
$maxPhoto = $Photos.Count
$starttime = Get-Date
$totalTime = 0;
#Copy loop
foreach ($photo in $Photos) {
    $tempdate = Get-Date $photo.LastWriteTime -Format 'dd.MM.yyyy'
    $photoyear = Get-Date $photo.LastWriteTime -Format 'yyyy'
    $filemonth = Get-Date $photo.LastWriteTime -Format 'MM'
    $monthfolder = $destfolder + $photoyear + '\кпп\' + $months[$filemonth - 1] + '\'
    $dayfolder = $monthfolder + $tempdate + '\'
    $destPhoto = $dayfolder + $photo.Name
    if (!(test-path $destPhoto)){
        if (!(test-path $dayfolder)){
            New-Item -ItemType Directory -Force -Path $dayfolder | Out-Null
        }
        Copy-Item -path $photo.FullName -Destination $dayfolder
        $reallycopied ++
    }
    
    #Count speed for progress bar each 10 files
    if ((($reallycopied % $speedmeasurepoint) -eq 0) -and ($reallycopied -gt 0)) {
        $elapsed = (Get-Date) - $starttime
        $speed = $reallycopied / $elapsed.TotalSeconds
        $remainsec = ($maxPhoto - $i) * $speed
	    #$totalTime = ($elapsed.TotalSeconds)*$maxPhoto / $reallycopied
        #$elapsed = (Get-Date) - $starttime
        #$remainsec = $totalTime - $elapsed.TotalSeconds
        $remainmin = [Math]::Floor($remainsec / 60)
        $remainsec = [Math]::Floor($remainsec % 60)
        if ($remainmin -lt 0) {
            $remainmin = 0
        }
        if ($remainsec -lt 0) {
            $remainsec = 0
        }

    }
    #Show progress bar
    $i++
    $percent = [math]::floor([int]$i*100/$maxPhoto)
    $statusmsg = 'Completed: ' + $percent + '%, Files processed: ' + $i + '/' + $maxPhoto
    if ($reallycopied -gt 0) {
        $statusmsg += ', Files copied: ' + $reallycopied
        if ($reallycopied -gt $speedmeasurepoint){
            $statusmsg += ', Time left: ' + $remainmin + ' m. ' + $remainsec + ' s'
        }
    }
    Write-Progress -Activity 'Copying files' -Status $statusmsg -PercentComplete $percent
}

if (!$isUSBconnected){
    #Disconnect from source
    net use /d $sourcefolder | Out-Null
}

#Print results
$skipped = $i - $reallycopied
$CR_LF = "`r`n"
$resultstring = '-----------------------------------------' + $CR_LF
$resultstring += 'Copying finished!' + $CR_LF
$resultstring += '--Totally copied: ' + $reallycopied + ' files.' + $CR_LF
$resultstring += '--Skipped: ' + $skipped + ' files' + $CR_LF
$resultstring += '-----------------------------------------' + $CR_LF
$resultstring += 'Press any key to exit...'
Read-Host -Prompt $resultstring