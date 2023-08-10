#Creating new directory to save output file
echo "Started the Automation"
$currentDate = Get-Date -Format "yyyyMMdd"
$OutputFiles = "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\$currentDate"
echo "$OutputFiles"
New-item $OutputFiles -itemType Directory
#$subFolder = New-Item -ItemType Directory -Path "$currentDate\$OutputFiles"

echo "collect the FILE"
#Taking source File 
$inputFile = "C:\Users\rohith.ss\Desktop\ReportWork\DataToExprt"       

$files = Get-ChildItem –Path $inputFile

$i=0
foreach($input in $files.FullName){
echo "$files"

$outputfile = $files[$i].BaseName
echo "BaseName $outputfile"
$excel = new-object -com excel.application
$workbook = $excel.workbooks.Add()
$worksheet =$workbook.Worksheets.Item(1)
$worksheet.name = $outputfile
$output = "$outputfile"+"Excel"+".xlsx"
echo "exporting"
$excel.DisplayAlerts = $False

$tempcsv = $excel.Workbooks.Open(“$($input)”)
$tempsheet = $tempcsv.Worksheets.Item(1)

$tempSheet.UsedRange.Copy() | Out-Null
$worksheet.Paste()
$tempcsv.close()
echo "Closed"
$range = $worksheet.UsedRange
$range.EntireColumn.Autofit() | out-null


$workbook.saveas("$OutputFiles\$output")
echo "Saved"

$excel.quit()
$i++

}
