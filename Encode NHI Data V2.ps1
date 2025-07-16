Start-Transcript -path C:\b2bex\ediclient\log\Process.log -Force -Append –NoClobber
$start = Get-Date

$input_folder = "C:\b2bex\ediclient\data\receive\"
$out_folder = "C:\b2bex\ediclient\data\output"
$filename_patten = "*.dat"

# 输出目录不存在就创建一个
If (-NOT (Test-path $out_folder)) { New-Item -Path $out_folder -ItemType Directory }

# 确保Archive目录存在
If (-NOT (Test-path "$input_folder\Archive")) { New-Item -Path "$input_folder\Archive" -ItemType Directory }

# 先备份所有文件到Archive目录
Get-ChildItem -Path $input_folder -File | Copy-Item -Destination $input_folder\Archive\ -Force

# Rename non-.dat files (excluding directories) with new naming rule
Get-ChildItem $input_folder -File | Where-Object { $_.Extension -ne ".dat" } | ForEach-Object {
    # Extract first 18 characters of base filename (without extension)
    $baseName = $_.BaseName
    if ($baseName.Length -gt 18) {
        $shortName = $baseName.Substring(0, 18)
    } else {
        $shortName = $baseName
    }
    
    # Handle potential duplicates by adding counter
    $counter = 1
    $newName = "$shortName.dat"
    while (Test-Path (Join-Path -Path $input_folder -ChildPath $newName)) {
        $newName = "$shortName($counter).dat"
        $counter++
    }
    
    $newFullName = Join-Path -Path $input_folder -ChildPath $newName
    Rename-Item -Path $_.FullName -NewName $newName -Force
    Write-Output "Renamed file from $($_.Name) to $newName"
}

if ((test-path "$input_folder\*.dat") -eq $FALSE) {
    write-host "No files to process!" 
    Send-MailMessage -From NHIAlert@bio-rad.com -To bruce_ge@bio-rad.com -SmtpServer smtp.bio-rad.com -Subject "NHI Error" -Body "No files to process."
}
else {
    $Input_Files = (Get-ChildItem $input_folder -Filter $filename_patten).FullName
    $i=0
    Foreach ($File in $Input_Files) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File)
        if ($baseName.Length -gt 17) {
            $shortBase = $baseName.Substring(0, 17)
        } else {
            $shortBase = $baseName
        }
        $out_filename = "$out_folder\nhi$shortBase.txt"
        $i++
        Write-Output "[$i/$($Input_Files.count)]`tNow Processing File `"$File`", `tOutfile `"$out_filename`"."
        $cmdstr="C:\b2bex\m2pc\m2pc.exe -i`"$File`" -o`"$out_filename`" -pC:\b2bex\m2pc\ebctoasc.txt"
        invoke-expression $cmdstr
    }
    
    #Middleware team will pickup from below path and send the files to /appl/E1D/INT/Inbound/OTC/OTC_IDD_0261_SAP/TBP (Test) or /appl/E1P/ENH/OTC/OTC_EDD_0477/NHI/TBP (Production)
    Copy-Item $out_folder\*.txt "D:\NHI_Data\Production\" -force    
    #Copy-Item $out_folder\*.txt "D:\NHI_Data\Test\" -force    
        
    Start-Sleep -Seconds 10        
    Move-Item $input_folder\*.dat $input_folder\Archive\ -force
    Move-Item $out_folder\*.txt $out_folder\Archive\ -force
        
    $ws = New-Object -ComObject WScript.Shell  
    $wsr = $ws.popup("Encode succeeded.",3,"Information",64)
}

$end = Get-Date
Write-Host -ForegroundColor Red ('Total Runtime: ' + ($end - $start).TotalSeconds)
stop-transcript