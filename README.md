# SysBox
SysAdmin Toolbox Forked From https://gist.github.com/noahpeltier/a7c2b5640aa7d0fad25c66efdcf72051

Original Script Updated to work with MahApps 2 and PowerShell Core.
Set $PSStyle.OutputRendering = 'Host' added due to additional characters added in-front of the results.
Removed Secondary Menu Item because it is obsolete nowdays, it's just commented. 
Added some additional fields (CPUSerial, HddSerial, Monitor Serial Numbers) which I find useful. 
Removed VNC button, I am not using it, just commented as well. 
Assigned Enter to Computer Search - it is mostly used button. 
SCCM Remote pointed to MECM Console Path. 
Small fixes were needed to make restart-process and few other functions work with PowerShell Core. 

Additional textboxes added showing the user running the script and running mode - user or admin. 
