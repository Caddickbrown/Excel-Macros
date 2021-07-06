'This script was used in the IFS ERP System originally for copying and pasting a value to multiple pages one after another. It could be utilised elsewhere'

Set WshShell = WScript.CreateObject("WScript.Shell")

'User defines the amount of orders that need to be copied down
ordercount=InputBox("How many orders do you have?","Copy Down","Whole Numbers Please")

If IsEmpty(ordercount) Then
	'Cancel Box clicked
	MsgBox "The Script has been cancelled"
Else
  'Select window
  WshShell.AppActivate "Lot Batch Master - Daniel Caddick-Brown @ IFS Applications 8 SP 1 - Live Database - IFS Applications"
  'Wait - just in case
  WScript.Sleep 500
  'Copy Original data
  WshShell.SendKeys "^{c}"
  'Define a variable
  Dim a
  a=0
  'Start Loop'
  Do until a>ordercount

    WScript.Sleep 1000
    'Paste Info
    WshShell.SendKeys "^{v}"
    WScript.Sleep 1000
    'Save Record
    WshShell.SendKeys "^{s}"
    WScript.Sleep 1000
    'Accept Dialog box
    WshShell.SendKeys "{Enter}"
    WScript.Sleep 1000
    'Load up next Record
    WshShell.SendKeys "^{DOWN}"
    'Add one to the count
    a=a+1
  Loop
End If
