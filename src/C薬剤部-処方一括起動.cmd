powershell.exe -executionpolicy bypass -command "& { . Y:\•a‰@\ˆã‹Ç\Y•wl‰È\‚â‚Ü‚à‚Æ\‰ğÍ\Tools-UserInteraction.ps1; . Y:\•a‰@\ˆã‹Ç\Y•wl‰È\‚â‚Ü‚à‚Æ\‰ğÍ\Tools-CheckLogin.ps1; if (-not (Invoke-checkLogin) ) { Invoke-ErrorDialog 'ƒƒOƒCƒ“‚µ‚Ä‚­‚¾‚³‚¢'; exit 1 }}"
if %ERRORLEVEL%==1 exit

start powershell.exe -windowstyle hidden -executionpolicy bypass -File Y:\•a‰@\ˆã‹Ç\Y•wl‰È\‚â‚Ü‚à‚Æ\‰ğÍ\Utility-Luncher.ps1 -JsonPath Y:\•a‰@\ˆã‹Ç\Y•wl‰È\‚â‚Ü‚à‚Æ\‰ğÍ\Config-Luncher-–òÜ•”ˆ•û.json -Title "–òÜ•”í’“(ˆ•û)"
