Attribute VB_Name = "FormMe"
'FormMe v1 By ShaG Mail me at ooshago0@aol.com
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub CenterForm(Fr As Form)
Fr.Left = (Screen.Width - Fr.Width) / 2
Fr.Top = (Screen.Height - Fr.Height) / 2
End Sub
Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub


Sub FormDance(Fr As Form)
' This makes a form dance across the screen
'You might have to Edit The Pause in this
'Its Funny when you put it in the MouseUp Function
Fr.Left = 5: Pause (0.01): Fr.Left = 400: Pause (0.01)
Fr.Left = 700: Pause (0.01): Fr.Left = 1000: Pause (0.01)
Fr.Left = 2000: Pause (0.01): Fr.Left = 3000: Pause (0.01)
Fr.Left = 4000: Pause (0.01): Fr.Left = 5000: Pause (0.01)
Fr.Left = 4000: Pause (0.01): Fr.Left = 3000: Pause (0.01)
Fr.Left = 2000: Pause (0.01): Fr.Left = 1000: Pause (0.01)
Fr.Left = 700: Pause (0.01): Fr.Left = 400: Pause (0.01)
Fr.Left = 5: Pause (0.01): Fr.Left = 400: Pause (0.01)
Fr.Left = 700: Pause (0.01): Fr.Left = 1000: Pause (0.01)
Fr.Left = 2000
End Sub
Sub GrafittiForm(frm As Form)
'This Is Real Phatty
'Got this at Dip!
If frm.WindowState = vbMinimized Then Exit Sub
frm.BackColor = vbBlack
frm.ScaleHeight = 100
frm.ScaleWidth = 100
For X = 0 To 300
DoEvents
X1 = Int(Rnd * 101)
X2 = Int(Rnd * 101)
Y1 = Int(Rnd * 101)
Y2 = Int(Rnd * 101)
colo = Int(Rnd * 15)
frm.Line (X1, Y1)-(X2, Y2), QBColor(colo)
frm.Line (X1, Y2)-(X2, Y1), QBColor(colo)
frm.Line (X2, Y1)-(X1, Y2), QBColor(colo)
frm.Line (Y1, Y2)-(X1, X2), QBColor(colo)
Next X
End Sub
Sub LoadText(thatxt As TextBox, File As String)
'I.e: Call LoadText(Text1,"C:\WINDOWS\filename.ext")
On Error GoTo Error
Dim mystr As String
Open File For Input As #1
Do While Not EOF(1)
Line Input #1, a$
texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
Loop
thatxt = texto$
Close #1
Exit Sub
Error:
X = MsgBox("Error Loading Text.. Sorry", vbOKOnly, "Error!!")
End Sub

Sub StayOnTop(frm As Form)
'I.e: StayOnTop Me
SetWinOnTop = SetWindowPos(frm.hwnd, -1, 0, 0, 0, 0, &H2 Or &H1)
End Sub
Sub ListToList(lst1 As ListBox, lst2 As ListBox)
For X% = 0 To SourceList.ListCount - 1
lst2.AddItem lst1.List(X%)
Next X%
End Sub
Function ListCount(thalst As ListBox) As Integer
'I.e : MsgBox Listcount(list1)
ListCount = thalst.ListCount - 0
End Function
Function ListDups(thalist As ListBox, Check) As Boolean
Dim X As Integer
For X = 0 To thalist.ListCount - 1
If (Check) = (thalist.List(X)) Then
ListDups = True
Exit Function
Else
End If
Next X
ListDups = False
End Function
Sub LoadList(thalst As ListBox, File As String)
'I.e:Call LoadList(List1, "C:\WINDOWS\filename.ext")
On Error GoTo Error
Open File For Input As #1
Do Until EOF(1)
Input #1, a$
thalst.AddItem a$
Loop
Close 1
Exit Sub
Error:
X = MsgBox("Sorry The List Wasnt Found!", vbOKOnly, "Error!!")
End Sub
Sub SaveList(thalst As ListBox, File As String)
'I.e:Call SaveList(List1, "C:\WINDOWS\filename.ext")
On Error GoTo Error
Open File For Output As #1
For i = 0 To thalst.ListCount - 1
a$ = thalst.List(i)
Print #1, a$
Next
Close 1
Exit Sub
Error:
X = MsgBox("Error savinmg list!", vbOKOnly, "Error!!")
End Sub


Function TextCharCount(thatxt As TextBox) As Integer
'I.e : MsgBox CharCount(Text1)
CharCount = Len(thatxt.Text)
End Function

Sub SaveText(thatxt As TextBox, File As String)
'I.e: Call SaveText(Text1,"C:\WINDOWS\filename.ext")
On Error GoTo Error
Dim mystr As String
Open File For Output As #1
Print #1, thatxt
Close 1
Exit Sub
Error:
X = MsgBox("Error Saving Text... Sorry", vbOKOnly, "Error!!")
End Sub
