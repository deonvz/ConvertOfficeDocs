# This script bulk converts Word office documents into PDF files. This is helpful for archival of files
# Example .doc to .pdf
# This script performs this conversion action for Word files to PDF in the given path directory. 
# Deon van Zyl
# Remember to create a folder C:\ConvertToPDF and to run this script as a administrator.
$path = "C:\ConvertToPDF" #Target directory for converting Word files
$word_app = New-Object -ComObject Word.Application

#Convert .doc and .docx to .pdf
Get-ChildItem -Path $path -Filter *.doc? | ForEach-Object {
    $document = $word_app.Documents.Open($_.FullName)
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
}
$word_app.Quit()