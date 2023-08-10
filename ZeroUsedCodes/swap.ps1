$currentDate = Get-Date -Format "yyyyMMdd"
$Sheets = Get-Item -Path "C:\Users\rohith.ss\Desktop\ReportWork\DataToExprt\ft_req_err_msg.csv" | Where-Object {$_.PSIsContainer -eq $false}
$excel = new-object -com excel.application

foreach ($Sheet in $Sheets) {
    $wb = $excel.Workbooks.Open($Sheet.FullName)
    $ws = $wb.ActiveSheet

    $c = $ws.Columns
    $c.Item(1).Cut()
    $c.Item(3).Insert()
    $wb.Close()
    }echo "Swapped the error block"