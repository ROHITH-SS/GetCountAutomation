$items= Get-ChildItem "C:\Users\rohith.ss\Desktop\ABCTEST"

$csv = Import-Csv "C:\Users\rohith.ss\Desktop\ABCTEST"

foreach($item in $items){
[pscustomobject]@{
oft_req_id=$item.oft_req_id
errmsg=$item.err_msg
result=$csv.Where({$item.oft_req_id -eq $_.oft_req_id}).result
}
}
echo result
