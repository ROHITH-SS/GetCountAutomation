$csvData = Import-Csv -Path "C:\Users\rohith.ss\Desktop\ReportWork\ft_req_err_msgExcel_updated.csv"

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




# Display the counts
Write-Host "NFT_EXTRACTED N: $countData1"
Write-Host "SUCCESS: $countData2"
Write-Host "Special Y: $countData3"
Write-Host "NFT-N: $countData4"
Write-Host "NFT-Y: $countData5"

$filteredData1 = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }
$countData10 = $filteredData1 | Group-Object -Property $columnName1 | Select-Object Count, Name

$filteredData2 = $csvData | Where-Object { $_.$columnName2 -eq "NEW" }
$countData11 = $filteredData2 | Group-Object -Property $columnName1 | Select-Object Count, Name

# Display the counts
Write-Host "Count of '$columnName1' with 'PROCESSING':"
$countData10 | Format-Table

Write-Host "Count of '$columnName1' with 'NEW':"
$countData11 | Format-Table



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






$columnName99 = "err_msg"
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
$countData2 = $filteredData | Group-Object -Property $columnName99 | Select-Object count , name
# Display the count with name
#Write-Host "Count of '$columnName99' with 'NFT_EXTRACTED' and '$columnName1' = 'N':"
$countData1 | Format-Table
$countData2 | Format-Table



$outputFilePath = "C:\Users\rohith.ss\Desktop\ReportWork\output.txt"
$output = $output1, $output2, $output3, $output4, $output5,$output6, $output7, $output8, $output9,$output10, $output11, $output12, $output13,  $output14, $output15  
                                                            
                                                            
$output1 ="NFT_EXTRACTED N: $countData1"
$output2 ="SUCCESS: $countData2"
$output3 ="Special Y: $countData3"
$output4 ="NFT-N: $countData4"
$output5 ="NFT-Y: $countData5"
$output6 =Write-Host "Count of '$columnName1' with 'PROCESSING':"
$output6 =$countData10 | Format-Table
$output7 =Write-Host "Count of '$columnName1' with 'NEW':"
$output7 =$countData11 | Format-Table
$output8 ="source_system"
$output8 =$countData 
$output9 =$countData2
$output10 =$rowCount
$output11 =$newReqId
$output12 =$processingReqId
$output13 =$lastRowData
$output14 =$countData1 
$output15 =$countData2                                                   
                                                            
                                                           
                                                           
                                                           
                                                           
                                                           
                                                           
$output | Out-File -FilePath $outputFilePath -Append              
                                                           
                                                           
                                                           