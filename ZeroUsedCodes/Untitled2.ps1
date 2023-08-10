$csvData = Import-Csv -Path "C:\Users\rohith.ss\Desktop\ReportWork\ft_req_err_msgExcel_updated.csv"

# Get the count of 'PROCESSING' column
$columnName = "request_status"
$processingCount = ($csvData | Where-Object { $_.$columnName -eq "PROCESSING" }).Count

# Get the count of 'special_request_ind' column
$columnName = "special_request_ind"
$specialRequestCount = ($csvData | Where-Object { $_.$columnName -eq "Y" }).Count

# Display the counts
Write-Host "Count of 'PROCESSING': $processingCount"
Write-Host "Count of 'special_request_ind' with value 'Y': $specialRequestCount"
