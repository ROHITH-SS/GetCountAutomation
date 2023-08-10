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
$output = "SUCCESS            : $countData2`n" +
          "Special Y          : $countData3`n" +
          "NFT_EXTRACTED N    : $countData1`n" +
          "NFT                : $countData41`n" +
          "NFT-Y              : $countData5`n" +
          "NFT-N              : $result`n"

# If Stuck in the Queue
$filteredData1 = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }
$countData10 = $filteredData1 | Group-Object -Property $columnName1 | Select-Object Count, Name

$filteredData2 = $csvData | Where-Object { $_.$columnName2 -eq "NEW" }
$countData11 = $filteredData2 | Group-Object -Property $columnName1 | Select-Object Count, Name

$output += "`nCount of '$columnName1' with 'PROCESSING':`n"
$output += $countData10 | Format-Table -AutoSize | Out-String
$output += "`nCount of '$columnName1' with 'NEW':`n"
$output += $countData11 | Format-Table -AutoSize | Out-String

$oftReqIdColumnName = "oft_req_id"

# Get the ID when request_status is NEW or PROCESSING
$newReqId = ($csvData | Where-Object { $_.$columnName2 -eq "NEW" }).$oftReqIdColumnName
$processingReqId = ($csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }).$oftReqIdColumnName

$output += "`n"
if ($newReqId) {
    $output += "ID(s) when request_status=NEW: $newReqId`n"
} else {
    $output += "No ID found when request_status=NEW`n"
}

if ($processingReqId) {
    $output += "ID(s) when request_status=PROCESSING: $processingReqId`n"
} else {
    $output += "No ID found when request_status=PROCESSING`n"
}

# source_system the counts
$columnName = "source_system"
$columnData = $csvData | Select-Object -ExpandProperty $columnName
$countData = $columnData | Group-Object | Select-Object Name, Count

$output += "`nsource_system"
$output += $countData | Format-Table -AutoSize | Out-String

$columnName11 = "request_type"
$columnData2 = $csvData | Select-Object -ExpandProperty $columnName11
$countData2 = $columnData2 | Group-Object | Select-Object Name, Count

$output += "`nrequest_type"
$output += $countData2 | Format-Table -AutoSize | Out-String

# Specify the column name to check
$columnName = "oft_req_id"  # Replace with the actual column name

# Filter the first row
$firstRow = $csvData[0]
# Get the value from the specified column in the last row
$lastRowValue = $csvData[-1].$columnName

$output += "`n"
if ($lastRowValue -eq "$lastRowValue") {
    $output += "The value in the last row of '$columnName' column is       : $lastRowValue.`n"
} else {
    $output += "The value in the last row of '$columnName' column is not $lastRowValue.`n"
}

# Get the count of rows
$rowCount = $csvData.Count
$output += "`nTotal Count: $rowCount`n"

# Display the last row data
$lastRowData = $csvData[-1]
$output += "`n$lastRowData`n"

$columnName99 = "err_msg"
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT_EXTRACTED" }
$countData1 = $filteredData | Group-Object -Property $columnName99 | Select-Object count, name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "NFT" }
$countData2 = $filteredData | Group-Object -Property $columnName99 | Select-Object count, name
$filteredData = $csvData | Where-Object { $_.$columnName1 -eq "N" -and $_.$columnName2 -eq "PROCESSING" }
$countData3 = $filteredData | Group-Object -Property $columnName99 | Select-Object count, name

$output += "`nCount of '$columnName99' with 'NFT_EXTRACTED' and '$columnName1' = 'N':`n"
$output += $countData1 | Format-Table -AutoSize | Out-String
$output += "`nCount of '$columnName99' with 'NFT':`n"
$output += $countData2 | Format-Table -AutoSize | Out-String
$output += "`nCount of '$columnName99' with 'PROCESSING':`n"
$output += $countData3 | Format-Table -AutoSize | Out-String

# Export the output to a Notepad file
$outputPath = "C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\output.txt"
$output | Out-File -FilePath $outputPath

Write-Host "Exported the output to: $outputPath"
