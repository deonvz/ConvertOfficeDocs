#This script bulk converts old office documents into new x version ones. It places the old document in the OLD directory when it has completed
#Example .doc to .docx
# This script performs this conversion action for Powerpoint, Word and Excel. 
# NOTE: Powerpoint conversions will open the application to be able to convert.
# Deon van Zyl
# Remember to create a folder C:\ConvertToNewOffice and to run this script as a administrator.
$source = "C:\ConvertToNewOffice"
$appX = New-Object -ComObject Excel.Application
$appW = New-Object -ComObject Word.Application
$appP = New-Object -ComObject PowerPoint.Application
$FormatX = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$FormatW = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
$FormatP = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
$searchX = Get-ChildItem -Path $source -Recurse -Include *.xls -Exclude *.xlsx
$searchW = Get-ChildItem -Path $source -Recurse -Include *.doc -Exclude *.docx
$searchP = Get-ChildItem -Path $source -Recurse -Include *.ppt -Exclude *.pptx

$searchX | ForEach-Object {
$document = $appX.Workbooks.Open($_.FullName)
$filename = "$($_.DirectoryName)\$($_.BaseName).xlsx"
$document.SaveAs([ref] $filename, [ref]$FormatX)
$document.Close()
$path = "$($_.DirectoryName)\$($_.Name)"
Mkdir -Force "$($_.DirectoryName)\old"
$destination = "$($_.DirectoryName)\old\$($_.Name)"
Move-Item -Path $path -Destination $destination -Force
}
$appX.Quit()

$searchW | ForEach-Object {
$document = $appW.Documents.Open($_.FullName)
$filename = "$($_.DirectoryName)\$($_.BaseName).docx"
$document.SaveAs([ref] $filename, [ref]$FormatW)
$document.Convert()
$document.Close()
$path = "$($_.DirectoryName)\$($_.Name)"
Mkdir -Force "$($_.DirectoryName)\old"
$destination = "$($_.DirectoryName)\old\$($_.Name)"
Move-Item -Path $path -Destination $destination -Force
}
$appW.Quit()

$searchP | ForEach-Object {
$document = $appP.Presentations.Open($_.FullName)
$filename = "$($_.DirectoryName)\$($_.BaseName).pptx"
$document.SaveAs([ref] $filename, [ref]$FormatP)
$document.Close()
$path = "$($_.DirectoryName)\$($_.Name)"
Mkdir -Force "$($_.DirectoryName)\old"
$destination = "$($_.DirectoryName)\old\$($_.Name)"
Move-Item -Path $path -Destination $destination -Force
}
$appP.Quit()