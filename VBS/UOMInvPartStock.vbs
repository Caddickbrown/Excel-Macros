'This would fill out the Unit of Measure for batches in IFS - it would bring up menus and select the latest UoM that had been picked. Could likely be used for similar situations'

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.AppActivate "Lot Batch Master - Daniel Caddick-Brown @ IFS Applications 8 SP 1 - Live Database - IFS Applications"

'This would be looped afterwards - never got to looping it, ended up copy and pasting'

WScript.Sleep 500
WshShell.SendKeys "{f8}"
WScript.Sleep 200
WshShell.SendKeys "{f3}"
WScript.Sleep 200
WshShell.SendKeys "{Enter}"
WScript.Sleep 200
WshShell.SendKeys "{Enter}"
WScript.Sleep 200
WshShell.SendKeys "^{s}"
WScript.Sleep 200
WshShell.SendKeys "^{DOWN}"
