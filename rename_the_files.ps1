$directoryPath = "file_location"

$fileList = Get-ChildItem $directoryPath

foreach ($file in $fileList) {
    $newFileName = $file.Name -replace "old_data", "new_data"
    $newFilePath = Join-Path -Path $directoryPath -ChildPath $newFileName
    Rename-Item -Path $file.FullName -NewName $newFileName
}

# "file_location" replacing the data due to the Company Private Information
