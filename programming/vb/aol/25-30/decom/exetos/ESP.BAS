' ESP.FRM
Option Explicit
Const c0076 = 2 ' &H2%
Const c007C = 12 ' &HC%
Const c0090 = 16 ' &H10%

Sub Check3D1_Click (Value As Integer)
Dim l0022 As Variant
l0022 = Check3D1.Value
If  l0022 = False Then Text1.Enabled = False
If  l0022 = True Then Text1.Enabled = True
End Sub

Sub Check3D2_Click (Value As Integer)
Dim l0034 As Variant
l0034 = Check3D2.Value
If  l0034 = False Then Text2.Enabled = False
If  l0034 = True Then Text2.Enabled = True
End Sub

Sub Check3D3_Click (Value As Integer)
Dim l0046 As Variant
l0046 = Check3D3.Value
If  l0046 = False Then Text3.Enabled = False
If  l0046 = True Then Text3.Enabled = True
End Sub

Sub Check3D4_Click (Value As Integer)
Dim l0058 As Variant
l0058 = Check3D4.Value
If  l0058 = False Then Text4.Enabled = False
If  l0058 = True Then Text4.Enabled = True
End Sub

Sub Command3D1_Click ()
Dim l0068 As Integer
Dim l006C As Integer
Dim l0070 As Integer
Dim l0072 As Integer
Dim l0074 As Integer
' Const c0076 = 2 ' &H2%
Dim l007A As Integer
' Const c007C = 12 ' &HC%
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Integer
Dim l0086 As Integer
Dim l0088 As Integer
' Const c0090 = 16 ' &H10%
Dim l0094 As Integer
Dim l0096 As Integer
Dim l0098 As Integer
Dim l009A As Integer
Dim l009C As Integer
Dim l009E As Integer
Dim l00A0 As Integer
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
If  Text1.Text = "" Then
If  Text2.Text = "" Then
If  Text3.Text = "" Then
If  Text4.Text = "" Then MsgBox "How Many Times To TOS?", 16, "error": Exit Sub
End If
End If
End If
If  Text1.Enabled = False Then
If  Text2.Enabled = False Then
If  Text3.Enabled = False Then
If  Text4.Enabled = False Then MsgBox "What TOS Methods?", 16, "error": Exit Sub
End If
End If
End If
If  Text1.Enabled = True Then
If  TOSer.Label7.Caption = "S" Then GoTo L28DA0
If  TOSer.Label7.Caption = "s" Then GoTo L29EE0
If  TOSer.Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  TOSer.Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
Exit Sub
L28DA0:
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L28E8E:
If  Text1.Text = "0" Then GoTo L29112
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text1.Text = Text1.Text - 1
 If  Label7.Caption = "s" Then Exit Sub
GoTo L28E8E
L29112:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
End If
If  Text2.Enabled = True Then
 If  Label7.Caption = "s" Then Exit Sub
Call sub870("GUIDEPAGER")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L29264:
If  Text2.Text = "0" Then GoTo L29508
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
 If  Label7.Caption = "s" Then Exit Sub
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
 If  Label7.Caption = "s" Then Exit Sub
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text2.Text = Text2.Text - 1
GoTo L29264
L29508:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
End If
If  Text3.Enabled = True Then
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
l0096% = extfn1B40(l0070%, c0076)
l0098% = extfn1B40(l0096%, c0076)
sub368 (l0098%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report A Violation")
Loop Until l0072% <> 0
SendKeys "{tab}" + "{ }"

L296F8:
If  Text3.Text = "0" Then GoTo L299E6
Do
DoEvents
l009A% = extfn1A98(l0068%, "Violations via Instant Messages")
Loop Until l009A% <> 0
l0074% = extfn1A98(l009A%, "Date")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, Date)
l0082% = extfn1A98(l009A%, "Time AM/PM")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, Time)
l009C% = extfn1A98(l009A%, "CUT and PASTE a copy of the IM here")
l009E% = extfn1B40(l009C%, c0076)
l00A0% = extfn19B8(l009E%, c007C, 0, TOSer.Combo1.Text + ":  " + TOSer.Text1.Text)
l0088% = extfn1A98(l009A%, "Send")
sub368 (l0088%)
subE20
Text3.Text = Text3.Text - 1
subD08 .7#
SendKeys "{ }"
GoTo L296F8
L299E6:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
l0094% = extfn1A28(l009A%, c0090, 0, 0&)
l0094% = extfn1A28(l0072%, c0090, 0, 0&)
End If
If  Text4.Enabled = True Then
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L29B64:
If  Text4.Text = "0" Then GoTo L29DE8
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text4.Text = Text4.Text - 1
 If  Label7.Caption = "s" Then Exit Sub
GoTo L29B64
L29DE8:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
End If

L29E1C:
Let l00A6 = TOSer.Combo2.List(l00A2)
l00A2 = l00A2 + 1
l00AA = Trim(l00A6)
If  l00AA = "" Then GoTo L29EA2
If  UCase$(l00AA) = UCase$(TOSer.Combo1) Then Exit Sub
GoTo L29E1C:
L29EA2:
TOSer.Combo2.AddItem TOSer.Combo1
Exit Sub

L29EE0:
TOSer.Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + TOSer.Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: E-Special TOS - Phrase " + TOSer.Label3 + "§]v^ñï")
If  Text1.Enabled = True Then
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L2A14C:
If  Text1.Text = "0" Then GoTo L2A3D0
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text1.Text = Text1.Text - 1
 If  Label7.Caption = "s" Then Exit Sub
GoTo L2A14C
L2A3D0:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
Call subBF0("ïñ^v[§Kids Only TOS Completed§]v^ñï")
subD08 1E-99#
End If
If  Text2.Enabled = True Then
 If  Label7.Caption = "s" Then Exit Sub
Call sub870("GUIDEPAGER")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L2A584:
If  Text2.Text = "0" Then GoTo L2A828
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
 If  Label7.Caption = "s" Then Exit Sub
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
 If  Label7.Caption = "s" Then Exit Sub
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text2.Text = Text2.Text - 1
GoTo L2A584
L2A828:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
Call subBF0("ïñ^v[§GuidePager TOS Completed§]v^ñï")
subD08 1E-99#
End If
If  Text3.Enabled = True Then
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
l0096% = extfn1B40(l0070%, c0076)
l0098% = extfn1B40(l0096%, c0076)
sub368 (l0098%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report A Violation")
Loop Until l0072% <> 0
SendKeys "{tab}" + "{ }"

L2AA7C:
If  Text3.Text = "0" Then GoTo L2AD6A
Do
DoEvents
l009A% = extfn1A98(l0068%, "Violations via Instant Messages")
Loop Until l009A% <> 0
l0074% = extfn1A98(l009A%, "Date")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, Date)
l0082% = extfn1A98(l009A%, "Time AM/PM")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, Time)
l009C% = extfn1A98(l009A%, "CUT and PASTE a copy of the IM here")
l009E% = extfn1B40(l009C%, c0076)
l00A0% = extfn19B8(l009E%, c007C, 0, TOSer.Combo1.Text + ":  " + TOSer.Text1.Text)
l0088% = extfn1A98(l009A%, "Send")
sub368 (l0088%)
subE20
Text3.Text = Text3.Text - 1
subD08 .7#
SendKeys "{ }"
GoTo L2AA7C
L2AD6A:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
l0094% = extfn1A28(l009A%, c0090, 0, 0&)
l0094% = extfn1A28(l0072%, c0090, 0, 0&)
Call subBF0("ïñ^v[§IM Violation [KW] TOS Completed§]v^ñï")
subD08 1E-99#
End If
If  Text4.Enabled = True Then
Call sub870("KO HELP")
l0068% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l006C% = extfn1A98(l0068%, "I Need Help!")
Loop Until l006C% <> 0
l0070% = extfn1AD0(l006C%, "_AOL_ICON")
L2AF52:
If  Text4.Text = "0" Then GoTo L2B1D6
sub368 (l0070%)
Do
DoEvents
l0072% = extfn1A98(l0068%, "Report Password Solicitations")
Loop Until l0072% <> 0
l0074% = extfn1A98(l0072%, "Screen Name of Member Soliciting Your Information:")
l007A% = extfn1B40(l0074%, c0076)
l0080% = extfn19B8(l007A%, c007C, 0, TOSer.Combo1.Text)
l0082% = extfn1A98(l0072%, "Copy and Paste the solicitation here:")
l0084% = extfn1B40(l0082%, c0076)
l0086% = extfn19B8(l0084%, c007C, 0, TOSer.Text1.Text)
l0088% = extfn1A98(l0072%, "Send")
sub368 (l0088%)
subE20
Text4.Text = Text4.Text - 1
 If  Label7.Caption = "s" Then Exit Sub
GoTo L2AF52
L2B1D6:
l0094% = extfn1A28(l006C%, c0090, 0, 0&)
Call subBF0("ïñ^v[§StaffPager [KW] TOS Completed§]v^ñï")
subD08 1E-99#
End If
Call subBF0("ïñ^v[§" + TOSer.Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + TOSer.Label5 + "§]v^ñï")
L2B330:
Let l00A6 = TOSer.Combo2.List(l00A2)
l00A2 = l00A2 + 1
l00AE = Trim(l00A6)
If  l00AE = "" Then GoTo L2B3B6
If  UCase$(l00AE) = UCase$(TOSer.Combo1) Then Exit Sub
GoTo L2B330:
L2B3B6:
TOSer.Combo2.AddItem TOSer.Combo1
End Sub


Sub Command3D2_Click ()
Me.Hide
End Sub

Sub Form_Load ()
subCD0 Me
End Sub
