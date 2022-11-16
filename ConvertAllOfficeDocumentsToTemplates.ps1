#This script bulk converts old office documents into new x version ones. It places the old document in the OLD directory when it has completed
#Example .doc to .docx
# This script performs this conversion action for Powerpoint, Word and Excel. 
# NOTE: Powerpoint conversions will open the application to be able to convert.
# Deon van Zyl
# Remember to create a folder C:\ConvertToTemplate and to run this script as a administrator.
$source = "C:\ConvertToTemplate"
$appX = New-Object -ComObject Excel.Application
$appW = New-Object -ComObject Word.Application
$appP = New-Object -ComObject PowerPoint.Application
$FormatX = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLTemplate
$FormatW = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLTemplate
$FormatP = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLTemplate
$searchX = Get-ChildItem -Path $source -Recurse -Include *.xlsx -Exclude *.xltx
$searchXM = Get-ChildItem -Path $source -Recurse -Include *.xlsm -Exclude *.xltm #Macro enabled
$searchW = Get-ChildItem -Path $source -Recurse -Include *.docx -Exclude *.dotx
$searchWM = Get-ChildItem -Path $source -Recurse -Include *.docm -Exclude *.dotm #macro enabled
$searchP = Get-ChildItem -Path $source -Recurse -Include *.pptx -Exclude *.potx
$searchPM = Get-ChildItem -Path $source -Recurse -Include *.pptm -Exclude *.potm #macro enabled

$searchX | ForEach-Object {
$document = $appX.Workbooks.Open($_.FullName)
$filename = "$($_.DirectoryName)\$($_.BaseName).xltx"
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
$filename = "$($_.DirectoryName)\$($_.BaseName).dotx"
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
$filename = "$($_.DirectoryName)\$($_.BaseName).potx"
$document.SaveAs([ref] $filename, [ref]$FormatP)
$document.Close()
$path = "$($_.DirectoryName)\$($_.Name)"
Mkdir -Force "$($_.DirectoryName)\old"
$destination = "$($_.DirectoryName)\old\$($_.Name)"
Move-Item -Path $path -Destination $destination -Force
}
$appP.Quit()