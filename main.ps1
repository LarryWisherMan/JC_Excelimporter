$RawDataFolder  = ".\Raw"
$ProcessedDataFolder = ".\Processed"

$rawDataFiles = Get-ChildItem -Path $RawDataFolder -File
$processedDataFiles = Get-ChildItem -Path $ProcessedDataFolder -File


###region Todo - Process raw data files
###endregion


###region import processed data files from csv

$processedData = foreach ($file in $processedDataFiles) {
    Import-Csv -Path $file.FullName
}

###endregion


