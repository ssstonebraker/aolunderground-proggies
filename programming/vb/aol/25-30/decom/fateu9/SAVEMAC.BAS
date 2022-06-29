' SAVEMAC.FRM
Option Explicit

Sub Check1_Click ()
If  Check1.Value = 0 Then
File1.Pattern = "*.mdr"
Text1.Text = "*.mdr"
Else
File1.Pattern = "*.mdr;*.spm"
Text1.Text = "*.mdr;*.spm"
End If
End Sub

Sub Command3D1_Click ()
Dim l0032 As Variant
Dim l0036 As Variant
Dim l003A As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0058 As String
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As String
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
If  Command3D1.Caption = "Save" Then
l0032 = Len(Text1.Text)
l0036 = Mid(Text1.Text, l0032 - 2, 3)
l003A = InStr(1, l0036, "mdr", 1)
If  l003A Then
l0044 = Dir1.Path + "\" + Text1.Text
l0048 = Mid(Dir1.Path, Len(Dir1.Path), 1)
If  l0048 = "\" Then
l0044 = Dir1.Path + Text1.Text
End If
On Error Resume Next
Open l0044 For Binary As #1
If  Err Then
MsgBox "Invalid file: *.mdr", 16
Exit Sub
End If
l004C = LOF(1)
Close #1
If  l004C > 0 Then
l0050 = MsgBox("File Already Exists. Do you wish to replace it?", 32 + 4)
If  l0050 = 6 Then
GoTo 10
Else
Exit Sub
End If
End If
10 :
On Error Resume Next
Open l0044 For Binary As #1
If  Err Then
MsgBox "Invalid file: *.mdr", 16
Exit Sub
End If
l0058$ = "" + Text1.Text + "" + " - Fate Macro Draw; " + drawman.Text1.Text
Put #1, 1, l0058$
Close #1
macsave.Hide
Else
MsgBox "error"
End If
Else
l005E = Dir1.Path + "\" + Text1.Text
l0062 = Mid(Dir1.Path, Len(Dir1.Path), 1)
If  l0062 = "\" Then
l005E = Dir1.Path + Text1.Text
End If
On Error Resume Next
Open l005E For Binary As #1
If  Err Then
MsgBox "Invalid *.mdr", 16
Exit Sub
End If
l0066$ = String(LOF(1), 0)
Get #1, 1, l0066$
Close #1
If  Check1.Value = 0 Then
30 :
l0068 = InStr(1, l0066$, " - Fate Macro Draw; ", 1)
If  l0068 Then
l006C = Mid(l0066$, l0068 + 20, Len(l0066$))
drawman.Text1.Text = l006C
l0070 = Mid(Text1.Text, 1, Len(Text1.Text) - 4)
drawman.Caption = "Macro Draw [" + l0070 + "]"
macsave.Hide
Else
MsgBox "Invalid *.mdr file", 16
Exit Sub
End If
Else
l0074 = InStr(1, l0066$, """", 1)
l0068 = InStr(1, l0066$, """" + "," + """", 1)
If  l0068 Then
l0070 = Mid(l0066$, 1, l0068 - 1)
l0070 = Mid(l0070, l0074 + 1, Len(l0070))
drawman.Caption = "Macro Draw " + "[" + l0070 + "]"
l006C = Mid(l0066$, l0068 + 3, Len(l0066$))
l0078 = InStr(l006C, """")
l006C = Mid(l006C, 1, l0078 - 1)
drawman.Text1.Text = l006C
macsave.Hide
Else
GoTo 30
Exit Sub
End If
End If
End If
End Sub

Sub Command3D2_Click ()
macsave.Hide
End Sub

Sub Dir1_Change ()
File1.Path = Dir1.Path
End Sub

Sub Drive1_Change ()
On Error Resume Next
Dir1.Path = Drive1.Drive
If  Err Then
MsgBox "Drive is not available", 16
Exit Sub
End If
End Sub

Sub File1_Click ()
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As String
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As String
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
l0082 = File1.List(File1.ListIndex)
Text1.Text = l0082
If  Command3D1.Caption = "Save" Then
l0082 = Len(Text1.Text)
l0086 = Mid(Text1.Text, l0082 - 2, 3)
l008A = InStr(1, l0086, "mdr", 1)
If  l008A Then
l008E = Dir1.Path + "\" + Text1.Text
l0092 = Mid(Dir1.Path, Len(Dir1.Path), 1)
If  l0092 = "\" Then
l008E = Dir1.Path + Text1.Text
End If
On Error Resume Next
Open l008E For Binary As #1
If  Err Then
MsgBox "Invalid file: *.mdr", 16
Exit Sub
End If
l0096 = LOF(1)
Close #1
If  l0096 > 0 Then
l009A = MsgBox("File Already Exists. Do you wish to replace it?", 32 + 4)
If  l009A = 6 Then
GoTo 1000
Else
Exit Sub
End If
End If
1000 :
On Error Resume Next
Open l008E For Binary As #1
If  Err Then
MsgBox "Invalid file: *.mdr", 16
Exit Sub
End If
l009E$ = "" + Text1.Text + "" + " - Fate Macro Draw; " + drawman.Text1.Text
Put #1, 1, l009E$
Close #1
macsave.Hide
Else
MsgBox "error"
End If
Else
l00A0 = Dir1.Path + "\" + Text1.Text
l00A4 = Mid(Dir1.Path, Len(Dir1.Path), 1)
If  l00A4 = "\" Then
l00A0 = Dir1.Path + Text1.Text
End If
On Error Resume Next
Open l00A0 For Binary As #1
If  Err Then
MsgBox "Invalid *.mdr", 16
Exit Sub
End If
l00A8$ = String(LOF(1), 0)
Get #1, 1, l00A8$
Close #1
If  Check1.Value = 0 Then
3000 :
l00AA = InStr(1, l00A8$, " - Fate Macro Draw; ", 1)
If  l00AA Then
l00AE = Mid(l00A8$, l00AA + 20, Len(l00A8$))
drawman.Text1.Text = l00AE
l00B2 = Mid(Text1.Text, 1, Len(Text1.Text) - 4)
drawman.Caption = "Macro Draw [" + l00B2 + "]"
macsave.Hide
Else
MsgBox "Invalid *.mdr file", 16
Exit Sub
End If
Else
l00B6 = InStr(1, l00A8$, """", 1)
l00AA = InStr(1, l00A8$, """" + "," + """", 1)
If  l00AA Then
l00B2 = Mid(l00A8$, 1, l00AA - 1)
l00B2 = Mid(l00B2, l00B6 + 1, Len(l00B2))
drawman.Caption = "Macro Draw " + "[" + l00B2 + "]"
l00AE = Mid(l00A8$, l00AA + 3, Len(l00A8$))
l00BA = InStr(l00AE, """")
l00AE = Mid(l00AE, 1, l00BA - 1)
drawman.Text1.Text = l00AE
macsave.Hide
Else
GoTo 3000
Exit Sub
End If
End If
End If
End Sub

Sub Form_Load ()
extsub1050 macsave.hWnd, -1, 0, 0, 0, 0, 3
End Sub
