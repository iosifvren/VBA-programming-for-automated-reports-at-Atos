$desktopPath = [Environment]::GetFolderPath("Desktop")

function CreateAndMoveExcelFiles($keyword, $destination) {
    $excelFiles = Get-ChildItem -Path $desktopPath -Filter "*$keyword*.xls*" -File

    if ($excelFiles) {
        $folderName = "Folder_$keyword"
        $folderPath = Join-Path -Path $desktopPath -ChildPath $folderName
        New-Item -Path $folderPath -ItemType Directory | Out-Null

        $excelFiles | Move-Item -Destination $folderPath | Out-Null

        Move-Item -Path $folderPath -Destination $destination
    }
}

CreateAndMoveExcelFiles "XXXX" "file_location"

# "XXXX" replacing the data due to the Company Private Information
# Same applies for the "file_location"
