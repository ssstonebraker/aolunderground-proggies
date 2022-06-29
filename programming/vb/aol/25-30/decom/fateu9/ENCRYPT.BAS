' ENCRYPT.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As Variant
Dim l005C As Variant
Dim l0060 As Variant
Dim l0064 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
If  Text1 = "" Then
MsgBox "You must type in something to encrypt!", 16
Exit Sub
End If
If  Text2 = "" Then
MsgBox "You must type in a keyword!", 16
Exit Sub
End If
enstatus.Caption = "Encryption Status"
enstatus.Show
Text2.Enabled = False
Text1.Enabled = False
Command3D1.Enabled = False
Command3D2.Enabled = False
Command3D3.Enabled = False
Command3D4.Enabled = False
Command3D5.Enabled = False
l0048 = 100 / Len(Text1)
l004C = 3038 / Len(Text1)
l0050 = 1
For l0054 = 1 To Len(Text1.Text)
l0058 = Mid(Text1.Text, l0054, 1)
l005C = Asc(l0058)
l0060 = Mid(Text2.Text, l0050, 1)
l0050 = l0050 + 1
If  l0050 > Len(Text2.Text) Then l0050 = 1
l0064 = Asc(l0060)
l0068 = l005C + l0064
If  l0068 > 255 Then
l006C = l0068 - 255
l0070 = Chr(l006C)
Else
l0070 = Chr(l0068)
End If
l0074 = l0074 + l0070
l0078 = l0078 + l004C
enstatus.Picture1.Line (x,y)(0, 0)-(l0078, 500), QBColor(1), BF
DoEvents
Next l0054
Text1.Text = l0074
enstatus.Picture1.Cls
Unload enstatus
Text2.Enabled = True
Text1.Enabled = True
Command3D1.Enabled = True
Command3D2.Enabled = True
Command3D3.Enabled = True
Command3D4.Enabled = True
Command3D5.Enabled = True
End Sub

Sub Command3D2_Click ()
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
enstatus.Caption = "Decryption Status"
If  Text1.Text = "" Then
MsgBox "You must enter in something to decrypt!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You must enter in a keyword!", 16
Exit Sub
End If
Text2.Enabled = False
Text1.Enabled = False
Command3D1.Enabled = False
Command3D2.Enabled = False
Command3D3.Enabled = False
Command3D4.Enabled = False
Command3D5.Enabled = False
enstatus.Show
On Error Resume Next
l007C = 3038 / Len(Text1)
l0080 = 1
For l0084 = 1 To Len(Text1.Text)
l0088 = Mid(Text1.Text, l0084, 1)
l008C = Asc(l0088)
l0090 = Mid(Text2.Text, l0080, 1)
l0080 = l0080 + 1
If  l0080 > Len(Text2.Text) Then l0080 = 1
l0094 = Asc(l0090)
l0098 = l008C - l0094
If  l0098 < 1 Then
l009C = l0098 + 255
l00A0 = Chr(l009C)
Else
l00A0 = Chr(l0098)
End If
l00A4 = l00A4 + l00A0
l00A8 = l00A8 + l007C
enstatus.Picture1.Line (x,y)(0, 0)-(l00A8, 500), QBColor(1), BF
DoEvents
Next l0084
Text1.Text = l00A4
enstatus.Picture1.Cls
Unload enstatus
Text2.Enabled = True
Text1.Enabled = True
Command3D1.Enabled = True
Command3D2.Enabled = True
Command3D3.Enabled = True
Command3D4.Enabled = True
Command3D5.Enabled = True
End Sub

Sub Command3D3_Click ()
Clipboard.Clear
Clipboard.SetText Text1
End Sub

Sub Command3D4_Click ()
Text1.Text = ""
End Sub

Sub Command3D5_Click ()
encrypt.WindowState = 1
End Sub

Sub sub7870 ()
End Sub

Sub Form_Load ()
extsub1050 encrypt.hWnd, -1, 0, 0, 0, 0, 3
Text2.Text = "Test"
Text1.Text = "To use this, type in the message here, and enter in a keyword that you will use with your friends to keep your letters private.An example: Just click on Encrypt and this message will be scrambled into Ascii. You can read the message again by clicking Decrypt, but you must have the SAME PASSWORD, which was when you Encrypted it. Reminder: The password is CASE-SENSITIVE, meaning that it must be exactly to be decrypt to the original message. For an example click on encrypt while the keyword is ""Test"" and then change the keyword to ""test"" rather than ""Test"" and click decrypt. The message will not come out as it was."
End Sub

Sub sub9980 ()
Unload encrypt
End Sub
