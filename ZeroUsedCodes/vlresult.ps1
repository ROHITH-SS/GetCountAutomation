$data = @{}
$errtable=Import-Excel "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_ord_reqExcel.xlsx" | ForEach-Object { $data[$_.oft_req_id]}

$actualtable=Import-Excel "C:\Users\rohith.ss\Desktop\ReportWork\ExportToExcelOutput\20230527\ft_req_err_msgExcel.xlsx" |  Select-Object *, @{n='err_msg';e={if ($data.Contains($_.oft_req_id)) {$data[$_.oft_req_id]} else {'N/A'}}}

