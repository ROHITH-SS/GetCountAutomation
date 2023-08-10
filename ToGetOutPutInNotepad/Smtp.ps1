
$From = "rohith.ss@accenture.com"
$To = "vinitha.ponnambalam@accenture.com"
$Cc = "rohith.ss@accenture.com"
$Attachment = "C:\Users\rohith.ss\Desktop\ReportAUTOMATION\ABCTEST\Guide.docx"
$Subject = "Hi Testing"
$Body = "<h2>Guys, look at these snippet of automation!</h2><br><br>"
$Body += "He   !"
$SMTPServer = "smtp.office365.com"
$SMTPPort = "25"
Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -Attachments $Attachment