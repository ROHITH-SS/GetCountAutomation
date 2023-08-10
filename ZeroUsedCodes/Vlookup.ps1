$data = @{}
$errtable = Import-Excel "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_ord_reqExcel.xlsx" |
    ForEach-Object { $data[$_.oft_req_id] = $_ }

$actualtable = Import-Excel "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_req_err_msgExcel.xlsx" |
    Select-Object *, @{n='err_msg';e={if ($data.Contains($_.oft_req_id)) {$data[$_.oft_req_id].err_msg} else {'N/A'}}}

$actualtable | Export-Excel -Path "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_req_err_msgExcel_updated.xlsx" -NoTypeInformation

# Count rows with err_msg
$countErrMsg = ($actualtable | Where-Object { $_.request_status -eq 'NFT_EXTRACTED' -and $_.special_request_ind -eq 'N' -and $_.err_msg -ne 'N/A' }).Count

# Display the count
Write-Host "Count of rows with err_msg when request_status is NFT_EXTRACTED and special_request_ind is N: $countErrMsg"
