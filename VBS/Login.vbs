'This script would Login to a system without me typing anything in

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.AppActivate "KCML Client" 'Selects the relevant window
WshShell.SendKeys "10.50.37.134" 'Fills in first field - which, in this case, would be an IP Address
WshShell.SendKeys "{TAB}" 'Moves to next section
WshShell.SendKeys "user.name" 'Types in username
WshShell.SendKeys "{TAB}" 'Moves to next section
WshShell.SendKeys "Password123" 'Types in password
WshShell.SendKeys "{TAB}" 'Moves to next section
WshShell.SendKeys "{DOWN 2}" 'Presses Down twice to select info from a dropdown
WshShell.SendKeys "{Enter}" 'Hits Login
WScript.Sleep 5000 'Waits for system to load
WshShell.SendKeys "{Enter}" 'Hits ok on a warning message that would pop up
