#Send email with log file as attachment

. .\Config.ps1

$DateNow = $((Get-Date).ToString('yyyy-MM-dd'))
$DateNowFull = Get-Date
$LogFile = "$LogFolder\Pelican_federated_model_build_log_{0}.csv" -f $DateStartedText

try{
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.importance = 2
    $mail.subject = "ERROR: Pelican Federated Model Build for $DateNow"
    $mail.body = "There is an error while running the Federated Model Build.`n`nBuild started on <$DateStarted> and finished on <$DateNowFull>"
    $mail.to = "lawrenerno.jinkim@exyte.net;janetjasintha.lopez@exyte.net"
    $mail.Attachments.Add($LogFile)
    WriteLog-Full "Sending email to : lawrenerno.jinkim@exyte.net and janetjasintha.lopez@exyte.net"
    $mail.Send()
    Start-Sleep 20
    $outlook.Quit()
    }

catch{
    $Exception = $_.Exception.Message
    WriteLog-Full "$Exception" -Type ERROR
    }