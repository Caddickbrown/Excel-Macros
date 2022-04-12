'This would fill out the Unit of Measure for batches in IFS - it would bring up menus and select the latest UoM that had been picked. Could likely be used for similar situations'

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.AppActivate "Lot Batch Master - Daniel Caddick-Brown @ IFS Applications 8 SP 1 - Live Database - IFS Applications"

'The following part is looped, change the "Do until" to however many loops you want to do

OrderCount=Int(InputBox("How many orders do you have?","# of Batches...","Whole Numbers Please"))
'If box is empty or has a qty of 0 then this will cancel the script
If OrderCount<1 Then
	'Cancel Box clicked
	MsgBox "Script Cancelled"

  Else
    a=0
'Start Loop'
    Do While a<OrderCount

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

      a=a+1
  Loop
End If
