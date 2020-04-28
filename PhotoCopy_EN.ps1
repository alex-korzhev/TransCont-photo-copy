<#--------- G L O B A L S ----------#>

#Put month's names into an array
$months = ("Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь")

<#--------- F U N C T I O N S ---------#>

#Returns a path for a photo based on it's date
Function Get-PathFromDate($photoDate, $destRootFolder) {
    $date = $photoDate.Split(".")
    $result = $destRootFolder + "\" + $date[2] + "\кпп\" + $months[$date[1] -1] + "\" + $photoDate
    return $result
}

#Returns a list of dates of all selected photos from camera.
Function Create-DateList($photos) {
    $dateList = New-Object System.Collections.Generic.List[String]
    foreach ($p in $photos) {
        $tmpDate = Get-Date $p.LastWriteTime -Format 'dd.MM.yyyy'
        if ($dateList -notcontains ($tmpDate)) {$dateList.Add($tmpDate)}
    }
    return $dateList
}

#Creates folders in the destination folder for each date in the list.
Function Create-Folders ($dateList, $destRootFolder) {
    foreach ($date in $dateList) {
        $tmpFolder = Get-PathFromDate $date $destRootFolder
        if (!(test-path $tmpFolder)) {New-Item -ItemType Directory -Force -Path $tmpFolder | Out-Null}
    }
}

<#
    Returns folder selected by user or quits the app if cancelled.
    Copied from:
    https://gist.github.com/IMJLA/1d570aa2bb5c30215c222e7a5e5078fd
#>
Function Get-Folder(){
    $AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
    $Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
    $OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $OpenFileDialog.AddExtension = $false
    $OpenFileDialog.CheckFileExists = $false
    $OpenFileDialog.DereferenceLinks = $true
    $OpenFileDialog.Filter = "Folders|`n"
    $OpenFileDialog.Multiselect = $false
    $OpenFileDialog.Title = "Select folder with photos"
    $OpenFileDialogType = $OpenFileDialog.GetType()
    $FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
    $IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null)
    $null = $OpenFileDialogType.GetMethod('OnBeforeVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$IFileDialog)
    [uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
    $FolderOptions = $OpenFileDialogType.GetMethod('get_Options',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null) -bor $PickFoldersOption
    $null = $FileDialogInterfaceType.GetMethod('SetOptions',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$FolderOptions)
    $VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName,'System.Windows.Forms.FileDialog+VistaDialogEvents',$false,0,$null,$OpenFileDialog,$null,$null).Unwrap()
    [uint32]$AdviceCookie = 0
    $AdvisoryParameters = @($VistaDialogEvent,$AdviceCookie)
    $AdviseResult = $FileDialogInterfaceType.GetMethod('Advise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdvisoryParameters)
    $AdviceCookie = $AdvisoryParameters[1]
    $Result = $FileDialogInterfaceType.GetMethod('Show',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,[System.IntPtr]::Zero)
    $null = $FileDialogInterfaceType.GetMethod('Unadvise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdviceCookie)
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $FileDialogInterfaceType.GetMethod('GetResult',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$null)
    } 
    if (!($OpenFileDialog.FileName)) {Read-Host -Prompt 'Folder selection cancelled. Press "Enter" to exit the program';Exit}
    return $OpenFileDialog.FileName
}


<# ---------B O D Y ---------#>


# Set encoding for error output
& cmd /c ver | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("cp866")

# Get path to photos
$sourcefolder = (Get-Folder + "\")

# Get a sync depth in days
$daysDepth = Read-Host -Prompt 'Sync for how many days? Press Enter for 31 day (1 month)'
if (!$daysDepth) {$daysDepth = 31}

# Show user a corresponding time period
$periodfrom = (Get-Date).AddDays(-$daysDepth).ToString('dd.MM.yyyy')
$periodto = Get-Date -Format 'dd.MM.yyyy' 
Write-Host ('Copying photos from ' + $periodfrom + ' to ' + $periodto)

# Set destination root folder
$destfolder = '\\B92-FILE-03\K\АТО\'

#Create folder if not exists
if (!(test-path $destfolder)){New-Item -ItemType Directory -Force -Path $destfolder | Out-Null}

# Select only photo and video
$extensions =  '.mp4','.jpg','.jpeg','.png','.bmp','.mov'
$extensions += '.MP4','.JPG','.JPEG','.PNG','.BMP','.MOV'

# Find files in selected period
$Photos = Get-ChildItem -Path $sourcefolder -File -Recurse | Where-Object {($_.LastWriteTime -gt (Get-Date).AddDays(-$daysDepth)) -and ($extensions.Contains($_.extension))}

# Create folder structure in selected rootfolder
Create-Folders (Create-DateList($Photos)) $destfolder

# Variables for progress bar
$i = 0
$reallycopied = 0
$maxPhoto = $Photos.Count
$starttime = Get-Date
$totalTime = 0
$copyErrors = 0

# Copy loop
foreach ($photo in $Photos) {
    # Set destination path for each file
    $tempfolder = Get-PathFromDate(Get-Date $photo.LastWriteTime -Format 'dd.MM.yyyy') $destfolder
    $tempfolder += ("\" + $photo.Name)

    # Skip file if it exists in the destination folder
    if (!(test-path $tempfolder)){
        
        # Try copying file, show an error if one occures
        try {Copy-Item -path $photo.FullName -Destination $tempfolder;$reallycopied++}
        catch [system.exception] {Write-Host ("Error occured while copying file " + $photo.FullName);$copyErrors++}
        finally {}
    }
    $i++

    # Count speed for progress bar after 10 files
    if ($reallycopied -gt 10) {
        $elapsed = (Get-Date) - $starttime
        $speed = $reallycopied / $elapsed.TotalSeconds
        $remainingTime = ($maxPhoto - $i) * $speed
        $remainingMin = [Math]::Floor($remainingTime / 60)
        $remainingSec = [Math]::Floor($remainingTime % 60)
        if ($remainingMin -lt 0) {$remainingMin = 0}
        if ($remainingSec -lt 0) {$remainingSec = 0}
    }

    # Show progress bar
    $percent = [math]::floor([int]$i*100/$maxPhoto)
    $statusmsg = 'Completed: ' + $percent + '%, Files processed: ' + $i + '/' + $maxPhoto
    if ($reallycopied -gt 0) {
        $statusmsg += ', Files copied: ' + $reallycopied
        if ($reallycopied -gt 0){
            $statusmsg += ', Time left: ' + $remainingMin + ' m. ' + $remainingSec + ' s'
        }
    }
    Write-Progress -Id 1 -Activity 'Copying files' -Status $statusmsg -PercentComplete $percent
}

# Close progress bar
Write-Progress -id 1 -Activity 'Copying files' -Completed

# Show results
$CR_LF = "`r`n"
$resultstring = '-----------------------------------------' + $CR_LF
$resultstring += 'Copying finished!' + $CR_LF
$resultstring += '--Totally copied: ' + $reallycopied + ' files.' + $CR_LF
$resultstring += '--Skipped: ' + ($i-$reallycopied) + ' files.' + $CR_LF
if ($copyErrors -gt 0) {$resultstring += '--Failed to copy: ' + ($copyErrors) + ' files.' + $CR_LF}
$resultstring += '-----------------------------------------' + $CR_LF
Write-Host $resultstring
Read-Host -Prompt 'Press any key to exit...'