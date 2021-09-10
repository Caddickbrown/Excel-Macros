Set WshShell = WScript.CreateObject("WScript.Shell")
 
Dim max, min, rand
max=InputBox("Maximum Value","Maximum Value","Maximum Value")
If IsEmpty(max) Then
MsgBox "The Script has been cancelled"
Else
min=InputBox("Minimum Value","Minimum Value","Minimum Value")
If IsEmpty(min) Then
MsgBox "The Script has been cancelled"
Else
Randomize
rand = Int((max-min+1)*Rnd+min)
MsgBox rand
End If
End If