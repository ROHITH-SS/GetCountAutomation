$filePath = Get-ChildItem -Path C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\DataToExprt\ft_req_err_msg.csv -Filter "powershell*" | Select-Object -First 1

if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Information Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}
if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Action Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}
if ($filePath) {
    $fileContent = Get-Content $filePath
    # Filter and remove data
    $filteredContent = $fileContent | Where-Object { $_ -notlike "*Warning Message*" }
    # Overwrite the file with the filtered content
    $filteredContent | Set-Content $filePath
}