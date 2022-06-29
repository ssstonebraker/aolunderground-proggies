Attribute VB_Name = "Module1"
Sub AOL_stuff(desiredw, parentw, output) 'output is a text slot or label ie. text1.text
Z = aolhwnd = parentw
hchild = GetWindow(parentw, GW_CHILD)
a = GetClass(hchild)
output = a
If output = desiredw Then
Z = aolhwnd = hchild
MsgBox hchild
Else
Do: DoEvents
hchild = GetWindow(hchild, GW_HWNDNEXT)
a = GetClass(hchild)
output = a
Loop Until output = desiredw
Z = aolhwnd = hchild
MsgBox hchild
End If
End Sub


