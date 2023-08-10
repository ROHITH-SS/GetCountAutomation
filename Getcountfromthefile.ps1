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

#$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
#$countData4 = $filteredData.Count

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
Write-Host "NFT-Y              : $countData5"
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


