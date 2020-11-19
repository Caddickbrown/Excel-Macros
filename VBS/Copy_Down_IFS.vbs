'This script was used in the IFS ERP System originally for copying and pasting a value to multiple pages one after another. It could be utilised elsewhere'

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.AppActivate "Lot Batch Master - Daniel Caddick-Brown @ IFS Applications 8 SP 1 - Live Database - IFS Applications"

WScript.Sleep 500
WshShell.SendKeys "^{c}"

'The following part could be looped - I never got into actually looping it, but would cut and paste it instead'

WScript.Sleep 1000
WshShell.SendKeys "^{v}"
WScript.Sleep 1000
WshShell.SendKeys "^{s}"
WScript.Sleep 1000
WshShell.SendKeys "{Enter}"
WScript.Sleep 1000
WshShell.SendKeys "^{DOWN}"
