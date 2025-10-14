########################################################################
# File Transformer Script                                              #
# This script transforms files to the new format like *.docx, *.xlsx,  #
#*.pptx                                                                #
########################################################################

#paths
#Change the path to your OneDrive
#Ensure you have the necessary permissions to access and modify files

#Set OneDrive Path
$onedrive_path = "C:\Users\test1\JG-NET\testsite - Documents"
$path = $onedrive_path

#Set Backup path
$backup = "C:\Backup"

#Create Log Files Folder
try {
    
    New-Item -ItemType Directory -Path $onedrive_path -Name "ConverterLogs" -ErrorAction Stop
    $logPath = $onedrive_path + "/" + "ConverterLogs"
}
catch {
    Write-Host "Error creating log directory"
    
}

#Create Log File
try {
    New-Item -ItemType File -Path $logPath -Name "Log_WordFiles.txt" -ErrorAction Stop
    New-Item -ItemType File -Path $logPath -Name "Log_ExcelFiles.txt" -ErrorAction Stop
    New-Item -ItemType File -Path $logPath -Name "Log_PowerPointFiles.txt" -ErrorAction Stop
    Write-Host "Log files created successfully"
}
catch {
    Write-Host "Error creating log files, they may already exist"
}

#Log file paths
$logWord = $logPath + "\Log_WordFiles.txt"
$logExcel = $logPath + "\Log_ExcelFiles.txt"
$logPowerPoint = $logPath + "\Log_PowerPointFiles.txt"



#Save file to new format
##Word files to *.docx
#Close Word
try {
    Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
}
catch {
    Write-Host "Word closed!"
}
$word_app = New-Object -ComObject Word.Application
$Format = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
Get-ChildItem -Path $path -Filter *.doc -Recurse -Force | 
ForEach-Object `
{
    $document = $word_app.Documents.Open($_.FullName)
    $docx_filename = "$($_.DirectoryName)\$($_.BaseName).docx"
    $doc_filename = "$($_.DirectoryName)\$($_.BaseName).doc"
    $document.SaveAs([ref] $docx_filename, [ref]$Format)
    $document.Close()
    Move-Item -Path $doc_filename -Destination $backup -Force -ErrorAction SilentlyContinue
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Converted: $doc_filename to $docx_filename"
    Add-Content -Path $logWord -Value $logEntry
}
$word_app.Quit()
$word_app = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
#Close Word
try {
    Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
}
catch {
    Write-Host "Word closed!"
}


##Excel files to *.xlsx
#Close Excel
try {
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
}
catch {
    Write-Host "Excel closed!"
}
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
$folderpath = $onedrive_path
$filetype ="*xls"
Get-ChildItem -Path $folderpath -Include $filetype -recurse | 
ForEach-Object `
{
	$path = ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
	"Converting $path"
	$workbook = $excel.workbooks.open($_.fullname)
    $path += ".xlsx"
    $workbook.saveas($path, $xlFixedFormat)
    $workbook.close()
    Move-Item -Path $_.FullName -Destination $backup -Force
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Converted: $($_.FullName) to $path"
    Add-Content -Path $logExcel -Value $logEntry
}
$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
#Close Excel
try {
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
}
catch {
    Write-Host "Excel closed!"
}

