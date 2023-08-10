$hash1 = @{}
$ft_ord_req = Import-Csv C:\Users\rohith.ss\Documents\ABCTEST\ft_req_err_msg.csv
$ft_ord_req | ForEach-Object {
    $hash1[$_.oft_req_id] = @($_.err_msg, 0)
}
$ft_req_err_msg = Import-Csv  C:\Users\rohith.ss\Documents\ABCTEST\ft_ord_req.csv

$ft_req_err_msg | ForEach-Object {
    if ($hash1.ContainsKey($_.oft_req_id)) {
        ($hash1[$_.oft_req_id])[1]++
        [PSCustomObject] @{

            oft_req_id = $_.oft_req_id

            err_msg = $hash1.($_.oft_req_id)[0]       
        }
    }   else{
            [PSCustomObject]@{
                oft_req_id = $_.oft_req_id
                err_msg = "?--->"
              
            }
}
}
$hash1.GetEnumerator() | ForEach-Object {
    if ($_.Value[1] -eq 0) {
        Write-Host "oft_req_id $($_.Value[0]) in ft_ord_req.csv is unmatched" -ForegroundColor Yellow
    }
}

$out | Export-Csv 'C:\Users\rohith.ss\Documents\ABCTEST\matched_output.csv' -NoTypeInformation

$path = “C:\Users\rohith.ss\Documents\ABCTEST\matched_output.csv”
$path1="C:\Users\rohith.ss\Documents\ABCTEST\ft_ord_req.csv"
$Excel = New-Object -ComObject excel.application
$Excel.visible = $false
$Workbook = $excel.Workbooks.open($path)
$Worksheet = $Workbook.WorkSheets.item(“matched_output”)
$worksheet.activate() 
$range = $WorkSheet.Range(“B1”).EntireColumn
$range.Copy() | out-null
$Workbook1 = $excel.Workbooks.open($path1)
$Worksheet1 = $Workbook1.Worksheets.item("ft_ord_req")
$Range1 = $Worksheet1.Range(“I1”)
$Worksheet1.Paste($Range1) 
$workbook1.Save() 
$Excel.Quit()
Remove-Variable -Name excel
[gc]::collect()
[gc]::WaitForPendingFinalizers()