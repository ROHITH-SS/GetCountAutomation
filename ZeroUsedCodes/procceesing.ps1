$csvData = Import-Csv -Path "C:\Users\rohith.ss\Desktop\ReportWork\ft_req_err_msgExcel_updated.csv"

# Define the column names
$columnName1 = "special_request_ind"
$columnName2 = "request_status"

# Filter the data based on the conditions
$filteredData1 = $csvData | Where-Object { $_.$columnName2 -eq "PROCESSING" }
$countData10 = $filteredData1 | Group-Object -Property $columnName1 | Select-Object Count, Name

$filteredData2 = $csvData | Where-Object { $_.$columnName2 -eq "NEW" }
$countData11 = $filteredData2 | Group-Object -Property $columnName1 | Select-Object Count, Name

# Display the counts
Write-Host "Count of '$columnName1' with 'PROCESSING':"
$countData10 | Format-Table

Write-Host "Count of '$columnName1' with 'NEW':"
$countData11 | Format-Table
