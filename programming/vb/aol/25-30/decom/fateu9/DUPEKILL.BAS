' DUPEKILL.FRM
Option Explicit

Sub sub9748 ()
End Sub

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l0026 As Variant
Dim l002C As Variant
Dim l0030 As String
Dim l0034 As Variant
Dim l0038 As Variant
Dim l003C As Variant
Dim l0042 As Variant
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As String
Dim l0058 As Variant
Dim l005C As Variant
Dim l0060 As Variant
Dim l0064 As Variant
Dim l0068 As Variant
Dim l0072 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As String
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
If  Combo1.Text = "New Mail" Then l0020 = "New Mail"
If  Combo1.Text = "Old Mail" Then l0020 = "Old Mail"
If  Combo1.Text = "Mail Sent" Then l0020 = "Outgoing Mail"
If  Combo1.Text = "Flash Mail" Then l0020 = "Incoming FlashMail"
l0026 = extfn1018("AOL FRAME25", 0&)
l002C = extfn10F8(l0026, "Welcome")
l0030$ = String(30, 0)
l0034 = extfnFA8(l002C, l0030$, 250)
If  l0034 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0038 = extfn1018("AOL FRAME25", 0&)
l003C = extfn10F8(l0038, l0020)
l0042 = extfn10C0(l003C, "_AOL_TREE")
If  l003C = 0 Then
If  l0042 = 0 Then
MsgBox "Please open your mailbox", 16
Exit Sub
End If
End If
l0048 = extfnFE0(l0042, 1036, 0, 0)
If  l0048 = 0 Then
MsgBox "An error occurred: No Mail", 16
Exit Sub
End If
On Error Resume Next
l004C = 3593 / l0048
l0050 = 0
Do Until 2 > 3
2 :
DoEvents
l0054$ = String(255, 0)
l0058 = extfn11A0(l0042, 1034, l0050, l0054$)
l005C = InStr(l0054$, Chr(9))
l0060 = Mid(l0054$, l005C + 1, Len(l0054$))
l0064 = InStr(l0060, Chr(9))
l0068 = Mid(l0060, l0064 + 1, Len(l0060))
Label1.Caption = l0068
l0072 = Mid(l0068, 1, 5)
If  l0072 = "" Then
Label1.Caption = Mid(l0068, 6, Len(l0068))
End If
DoEvents
If  Label1.Caption = "" Then GoTo 3
l007C = InStr(1, text2.Text, Label1.Caption, 1)
If  l007C Then
l0080 = extfnFE0(l0042, 1031, l0050, 0)
l0084 = extfn10F8(l003C, "Delete")
l008A = fn1480(l0084)
DoEvents
l008E = extfn10F8(l0038, l0020)
l0092$ = String(30, 0)
l0094 = extfnFA8(l008E, l0092$, 250)
If  l0092$ Like "*FlashMail*" Then
DoEvents
890 :
DoEvents
l0098 = extfn1018("#32770", "America Online")
If  l0098 = 0 Then GoTo 890
l009C = extfn10F8(l0098, "&Yes")
l00A0 = fn1480(l009C)
End If
l00A4 = l00A4 + 1
l00A8 = l00A8 + l004C
Picture1.Line (x,y)(0, 0)-(l00A8, 500), QBColor(1), BF
GoTo 2
Else
text2.Text = text2.Text + "," + Label1.Caption
End If
l0050 = l0050 + 1
l00A8 = l00A8 + l004C
Picture1.Line (x,y)(0, 0)-(l00A8, 500), QBColor(1), BF
If  Err Then
MsgBox "Some error occured. If you do not have Windows 95, please purchase it for optimum performance!", 16
Exit Sub
End If
DoEvents
Loop
3 :
If  l00A4 = "" Then l00A4 = 0
MsgBox l00A4 + " duplicate mails deleted!"
text2.Text = ""
Label1.Caption = ""
Picture1.Cls
End Sub

Sub Command3D2_Click ()
help1.Show
End Sub

Sub Form_Load ()
Combo1.AddItem "New Mail"
Combo1.AddItem "Old Mail"
Combo1.AddItem "Mail Sent"
Combo1.AddItem "Flash Mail"
Combo1.ListIndex = 0
extsub1050 DUPEKILL.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub sub9980 ()
Unload DUPEKILL
End Sub
