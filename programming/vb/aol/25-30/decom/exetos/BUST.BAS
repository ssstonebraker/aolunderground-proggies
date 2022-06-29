' BUST.FRM
Option Explicit
Const c007C = 12 ' &HC%

Sub Combo1_KeyPress (KeyAscii As Integer)
Dim l001E As Integer
Dim l0022 As Integer
Dim l0026 As Variant
Dim l002A As String
Dim l002E As String
Dim l0036 As String
Dim l0038 As String
If  KeyAscii = 13 Then
l001E% = extfn1A60("aol frame25", 0&)
l0022% = fn528()
l0026 = fnE58(l0022%)
l002A$ = UCase$(l0026)
l002E$ = fnD78(l002A$)
l0036$ = UCase$(Combo1.Text)
l0038$ = fnD78(l0036$)
If  Combo1.Text = "" Then MsgBox "What Room You Gonna Bust into?", 16, "Error": Exit Sub
If  l0038$ = l002E$ Then MsgBox "You Are Already in " + Combo1.Text + "!", 16, "Error": Exit Sub
Label2.Caption = "0"
Label3.Caption = ""
Label4.Caption = "0"
Command3D2.Enabled = True
Command3D1.Enabled = False
Command3D3.Enabled = False
Combo1.Enabled = False
Command1.Value = True
End If
End Sub

Sub Command1_Click ()
Dim l0064 As Integer
Dim l0066 As Integer
Dim l0068 As Variant
Dim l006C As String
Dim l006E As String
Dim l0070 As String
Dim l0072 As String
Dim l0076 As Integer
Dim l007A As Integer
' Const c007C = 12 ' &HC%
Dim l0080 As Variant
Dim l0084 As Integer
Dim l0086 As Integer
L26526:
Label2.Caption = Label2.Caption + 1
l0064% = extfn1A60("aol frame25", 0&)
l0066% = fn528()
l0068 = fnE58(l0066%)
l006C$ = UCase$(l0068)
l006E$ = fnD78(l006C$)
l0070$ = UCase$(Combo1.Text)
l0072$ = fnD78(l0070$)
Combo1.Text = fnD78(Combo1)
If  l0072$ = l006E$ = True Then GoTo L26B4C
If  Label3.Caption = "S" Then
Command3D2.Enabled = False
Command3D1.Enabled = True
Command3D3.Enabled = True
Combo1.Enabled = True
Exit Sub
End If
l0064% = extfn1A60("aol frame25", 0&)
Call sub790("Keyword...")
Do
DoEvents
l0076% = extfn1A98(l0064%, "Keyword")
Loop Until l0076% <> 0
l007A% = extfn1AD0(l0076%, "_aol_edit")
l0080 = extfn19B8(l007A%, c007C, 0, "aol://2719:2-2-" + Combo1.Text)
l0084% = extfn1AD0(l0076%, "_aol_icon")
sub368 (l0084%)
subD08 .3#
Do
DoEvents
l0064% = extfn1A60("aol frame25", 0&)
l0066% = fn528()
l0068 = fnE58(l0066%)
l006C$ = UCase$(l0068)
l006E$ = fnD78(l006C$)
l0070$ = UCase$(Combo1.Text)
l0072$ = fnD78(l0070$)
Combo1.Text = fnD78(Combo1)
If  l0072$ = l006E$ = True Then GoTo L26B4C
l0072$ = fnD78(l0070$)
l006E$ = fnD78(l006C$)
l0086% = extfn1A60("#32770", "America Online")
If  l0072$ = l006E$ <> 0 Then GoTo L26B4C
If  l0086% <> 0 Then
subE20
If  l0072$ = l006E$ = True Then GoTo L26B4C
GoTo L26526:
End If
If  l0072$ = l006E$ = True Then GoTo L26B4C
Loop
If  l0086% <> 0 Then
subE20
GoTo L26526:
End If
l0064% = extfn1A60("aol frame25", 0&)
l0066% = fn528()
l0068 = fnE58(l0066%)
l006C$ = UCase$(l0068)
l006E$ = fnD78(l006C$)
l0070$ = UCase$(Combo1.Text)
l0072$ = fnD78(l0070$)
If  l0072$ = l006E$ = True Then GoTo L26B4C
Exit Sub
L26B4C:
subD08 .5#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§ﬂusted in " + Combo1 + " - " + Label2 + " Try(s)§]v^ñï")
Label2.Caption = "0"
Combo1.Enabled = True
Command3D2.Enabled = False
Command3D3.Enabled = True
Command3D1.Enabled = True
Label3.Caption = "S"
Exit Sub
End Sub

Sub Command3D1_Click ()
Dim l0088 As Integer
Dim l008A As Integer
Dim l008C As Variant
Dim l0090 As String
Dim l0092 As String
Dim l0094 As String
Dim l0096 As String
l0088% = extfn1A60("aol frame25", 0&)
l008A% = fn528()
l008C = fnE58(l008A%)
l0090$ = UCase$(l008C)
l0092$ = fnD78(l0090$)
l0094$ = UCase$(Combo1.Text)
l0096$ = fnD78(l0094$)
If  Combo1.Text = "" Then MsgBox "What Room You Gonna Bust into?", 16, "Error": Exit Sub
If  l0096$ = l0092$ Then MsgBox "You Are Already in " + Combo1.Text + "!", 16, "Error": Exit Sub
Label2.Caption = "0"
Label3.Caption = ""
Label4.Caption = "0"
Command3D2.Enabled = True
Command3D1.Enabled = False
Command3D3.Enabled = False
Combo1.Enabled = False
Command1.Value = True
End Sub

Sub Command3D2_Click ()
Label3.Caption = "S"
Command3D2.Enabled = False
Command3D1.Enabled = True
Command3D3.Enabled = True
Combo1.Enabled = True
End Sub

Sub Command3D3_Click ()
Me.Hide
End Sub

Sub Command3D4_Click ()
Me.WindowState = 1
End Sub

Sub Form_Load ()
subCD0 Me
Combo1.AddItem "Coldice"
Combo1.AddItem "DieV"
Combo1.AddItem "FreeShit"
End Sub

Sub sub36C0 ()
End Sub
