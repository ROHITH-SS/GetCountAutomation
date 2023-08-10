$csvData = Import-Csv -Path "C:\Users\rohith.ss\Desktop\ReportWork\ft_req_err_msgExcel_updated.csv"

# Filter the first row
$firstRow = $csvData[0]

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

# Get the value from the specified column in the last row
$lastRowValue = $csvData[-1].$columnName

# Check the value
if ($lastRowValue -eq "$lastRowValue") {
    Write-Host "The value in the last row of '$columnName' column is $lastRowValue."
} else {
    Write-Host "The value in the last row of '$columnName' column is not $lastRowValue."
}


# Get the count of rows
$rowCount = $csvData.Count

# Get the count of columns
$columnCount = $csvData[0].PSObject.Properties.Count

# Display the counts
Write-Host "Total Count: $rowCount"
#Write-Host "Column Count: $columnCount"



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

# Display the last row data
$lastRowData = $csvData[-1]
echo $lastRowData




# Define the column names
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "SUCCESS" }
$countData2 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "Y" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData3 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NFT" -and $_.$columnName1 -eq "N" }
$countData4 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NFT" -and $_.$columnName1 -eq "Y" }
$countData5 = $filteredData.Count

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" -and $_.$columnName1 -eq "Y" }
$countData6 = $filteredData.Count
Write-Host "PROCESSING-Y: $countData6"


$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NEW" -and $_.$columnName1 -eq "Y" }
$countData8 = $filteredData.Count
Write-Host "NEW-Y: $countData8"

$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "NEW" -and $_.$columnName1 -eq "N" }
$countData9 = $filteredData.Count
Write-Host "NEW-N: $countData9"


$filteredData = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" -and $_.$columnName1 -eq "N" }
$countData7 = $filteredData.Count
Write-Host "PROCESSING-N: $countData7"

# Display the counts
Write-Host "NFT_EXTRACTED N: $countData1"
Write-Host "SUCCESS: $countData2"
Write-Host "Special Y: $countData3"
Write-Host "NFT-N: $countData4"
Write-Host "NFT-Y: $countData5"







$columnName99 = "err_msg"
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
$countData2 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
# Display the count with name
Write-Host "Count of '$columnName99' with 'NFT_EXTRACTED' and '$columnName1' = 'N':"
$countData1 | Format-Table
$countData2 | Format-Table

