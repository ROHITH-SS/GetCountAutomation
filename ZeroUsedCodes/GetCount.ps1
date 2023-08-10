                   # Install the ImportExcel module if not already installed
#Install-Module -Name ImportExcel -Scope CurrentUser

# Import the Excel file
$excelData = Import-Excel -Path "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_ord_reqExcel.xlsx" -WorksheetName "ft_ord_req"
# Filter the first row
$firstRow = $excelData[0]

$columnName = "source_system"  
$columnData = $excelData[0..($excelData.Count - 0)] | Select-Object -ExpandProperty $columnName
$countData = $columnData | Group-Object | Select-Object Name, Count
$countData 



$columnName11 = "request_type"  
$columnData2 = $excelData[0..($excelData.Count - 0)] | Select-Object -ExpandProperty $columnName11
$countData2 = $columnData2 | Group-Object | Select-Object Name, Count
echo "
request_type
" $countData2




