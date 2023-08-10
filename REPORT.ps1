$filePath = Get-ChildItem -Path C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_req_err_msg.csv -Filter "powershell*" | Select-Object -First 1

if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Information Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}
if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Action Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}
if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Warning Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}
echo "UnWanted Data Filtered"

#VLOOKUP_METHOD
echo "VLOOKUP STARTED"
$hash1 = @{}
$ft_ord_req = Import-Csv C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_req_err_msg.csv
$ft_ord_req | ForEach-Object {
    $hash1[$_.oft_req_id] = @($_.err_msg, 0)
}
$ft_req_err_msg = Import-Csv  C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_ord_req.csv

  $output=$ft_req_err_msg | ForEach-Object {
    if ($hash1.ContainsKey($_.oft_req_id)) {
        ($hash1[$_.oft_req_id])[1]++
        [PSCustomObject] @{

            oft_req_id = $_.oft_req_id

            err_msg = $hash1.($_.oft_req_id)[0] 
                  
        }
    }   else{
            [PSCustomObject]@{
                oft_req_id = $_.oft_req_id
                err_msg = "NA"
              
            }
}
}
$hash1.GetEnumerator() | ForEach-Object {
    if ($_.Value[1] -eq 0) {
        Write-Host "oft_req_id $($_.Value[0]) in ft_ord_req.csv is unmatched" -ForegroundColor Yellow
    }
}
 

  $output | Export-Csv 'C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\matched_output1.csv' -NoTypeInformation

$path200 = "C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\matched_output1.csv"
$path1= "C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_ord_req.csv"

$Excel = New-Object -ComObject excel.application
$Excel.visible = $false
$Workbook = $Excel.Workbooks.open($path200)
$Worksheet = $Workbook.WorkSheets.item(“matched_output1”)
$worksheet.activate() 
$range = $WorkSheet.Range(“B1”).EntireColumn
$range.Copy() | out-null

#Specify the sheet name if you have differen sheet name
$Workbook1 = $Excel.Workbooks.open($path1)
$Worksheet1 = $Workbook1.Worksheets.item("ft_ord_req")

#Choose the column where the VLOOKUP have to save
$Range1 = $Worksheet1.Range(“I1”)
$Worksheet1.Paste($Range1) 
$Workbook1.Save() 
$Workbook1.Close()
$Workbook.Close()
$Excel.Quit()
Remove-Variable -Name excel
[gc]::collect()
[gc]::WaitForPendingFinalizers()

echo "VLOOKUP OVER"
#GetCountFromTheData
echo "________________________________________________________________________________________________________________________________"
echo "                                                      Count from the file"
$csvData = Import-Csv -Path "C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_ord_req.csv"

# Define the Excel Sheet column names
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions( to get count)
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "SUCCESS" }
$countData2 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "Y" }
$countData3 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
$countData4 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NFT" -and $_.$columnName1 -eq "Y" }
$countData5 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NFT" }
$countData41 = $filteredData.Count

$result = $countData41 - $countData5
# Display the counts
Write-Host "SUCCESS            : $countData2"
Write-Host "Special Y          : $countData3"
#Write-Host "NFT-N =Do substraction with : $result"
Write-Host "NFT_EXTRACTED N    : $countData1"
Write-Host "NFT                : $countData41"
#Write-Host "NFT-Y              : $countData5"
Write-Host "NFT-N              : $result"


#If Stuck in the Queue
$filteredData1 = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }
$countData10 = $filteredData1 | Group-Object -Property $columnName1 | Select-Object Count, Name

$filteredData2 = $csvData | Where-Object { $_.$columnName2 -eq "NEW" }
$countData11 = $filteredData2 | Group-Object -Property $columnName1 | Select-Object Count, Name
# Display the counts
Write-Host "Count of '$columnName1' with 'PROCESSING':"
$countData10 | Format-Table
Write-Host "Count of '$columnName1' with 'NEW':"
$countData11 | Format-Table

$oftReqIdColumnName = "oft_req_id"

# Get the ID when request_status is NEW or PROCESSING
$newReqId = ($csvData | Where-Object { $_.$columnName2 -eq "NEW" }).$oftReqIdColumnName
$processingReqId = ($csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }).$oftReqIdColumnName


# Display the IDs
if ($newReqId) {
    Write-Host "ID(s) when request_status=NEW: $newReqId"
} else {
    Write-Host "No ID found when request_status=NEW"
}

if ($processingReqId) {
    Write-Host "ID(s) when request_status=PROCESSING: $processingReqId"
} else {
    Write-Host "No ID found when request_status=PROCESSING"
}




#source_system  the counts
$columnName = "source_system"  
$columnData = $csvData | Select-Object -ExpandProperty $columnName
$countData = $columnData | Group-Object | Select-Object Name, Count
$countData

$columnName11 = "request_type"  
$columnData2 = $csvData | Select-Object -ExpandProperty $columnName11
$countData2 = $columnData2 | Group-Object | Select-Object Name, Count
echo "
request_type
" $countData2

# Specify the column name to check
$columnName = "oft_req_id"  # Replace with the actual column name

# Filter the first row
$firstRow = $csvData[0]
# Get the value from the specified column in the last row
$lastRowValue = $csvData[-1].$columnName

# Check the value
if ($lastRowValue -eq "$lastRowValue") {
    Write-Host "The value in the last row of '$columnName' column is       : $lastRowValue."
} else {
    Write-Host "The value in the last row of '$columnName' column is not $lastRowValue."
}


# Get the count of rows
$rowCount = $csvData.Count
# Display the counts
Write-Host "Total Count: $rowCount"

# Get the count of columns
#$columnCount = $csvData[0].PSObject.Properties.Count
#Write-Host "Column Count: $columnCount"

# Display the last row data
$lastRowData = $csvData[-1]
echo $lastRowData


$columnName99 = "err_msg"
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
$countData2 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "PROCESSING" }
$countData3 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
# Display the count with name
Write-Host "Count of '$columnName99' with 'NFT_EXTRACTED' and '$columnName1' = 'N':"
$countData1 | Format-Table
$countData2 | Format-Table
$countData3 | Format-Table


