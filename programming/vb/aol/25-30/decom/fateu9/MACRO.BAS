' MACRO.FRM
Option Explicit

Sub acs_Click ()
set1.Show
End Sub

Sub backc_Click ()
colors.Caption = "BackColor Selection"
colors.Show
End Sub

Sub Command3D1_Click ()
Dim l003A As Variant
Dim l0040 As Variant
Dim l0044 As String
Dim l0048 As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As Variant
Dim l0070 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0084 As Variant
Dim l008A As Variant
Dim l0090 As Variant
Dim l0096 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
If  Command3D1.Caption = "Scroll" Then
l003A = extfn1018("AOL FRAME25", 0&)
l0040 = extfn10F8(l003A, "Welcome")
l0044$ = String(30, 0)
l0048 = extfnFA8(l0040, l0044$, 250)
If  l0048 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1 = "" Then
MsgBox "You Must Enter In Something To Scroll", 16
Exit Sub
End If
Command3D1.Caption = "Stop"
For l0052 = 1 To Len(Text1.Text)
l0056 = Mid(Text1.Text, l0052, 1)
If  l0056 = Chr(32) Then
l005A = l005A + " "
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
GoTo 39
End If
l005A = l005A + l0056
39 :
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
If  l0056 = Chr(10) Then
l005E = Mid(l005A, 1, Len(l005A) - 2)
l0062 = extfn1018("AOL FRAME25", 0&)
l0066 = extfn10F8(l0062, "AOL Live!")
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
If  Check3D1.Value = True Then GoTo 50
If  l0070 = 4 Then
l0070 = 0
l007A = Label1.Caption / 10
l007E = Timer: While Timer - l007E < l007A: DoEvents: Wend
End If
50 :
l0084 = extfn1088(l0066)
l008A = extfn10C0(l0084, "_AOL_EDIT")
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
l0090 = fnA68(l005E)
l0070 = l0070 + 1
l0096 = extfn1168(l008A, 2)
l009C = fn1480(l0096)
DoEvents
l005A = ""
End If
If  l0052 = Len(Text1.Text) Then
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
If  Check3D1.Value = True Then GoTo 30
If  l0070 = 4 Then
l0070 = 0
l00A0 = Timer: While Timer - l00A0 < 2: DoEvents: Wend
End If
30 :
l00A4 = extfn1018("AOL FRAME25", 0&)
l00A8 = extfn10F8(l00A4, "AOL Live!")
If  l005A = "" Then Exit Sub
l00AC = extfn1088(l00A8)
l00B0 = extfn10C0(l00AC, "_AOL_EDIT")
l00B4 = fnA68(l005A)
l0070 = l0070 + 1
l00B8 = extfn1168(l00B0, 2)
l00BC = fn1480(l00B8)
DoEvents
l005A = ""
Command3D1.Caption = "Scroll"
Exit Sub
End If
If  Command3D1.Caption = "Scroll" Then GoTo L305A0
Next l0052
Else
Command3D1.Caption = "Scroll"
End If

L305A0:
Command3D1.Caption = "Scroll"
End Sub

Sub Command3D2_Click ()
Dim l00C0 As Variant
l00C0 = MsgBox("Are you sure you want to clear that?", 32 + 4)
If  l00C0 = 6 Then
Text1.Text = ""
Else
End If
End Sub

Sub Command3D3_Click ()
macsave.Show
macsave.Caption = "Save as *.mdr"
macsave.Command3D1.Caption = "Save"
End Sub

Sub Command3D4_Click ()
macsave.Show
macsave.Caption = "Load *.mdr"
macsave.Command3D1.Caption = "&Load"
End Sub

Sub copya_Click ()
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
l00C8 = Text1.SelStart
l00CC = Text1.SelLength
l00D0 = Mid(Text1.Text, 1, l00C8)
l00D4 = Mid(Text1.Text, l00C8 + l00CC + 1, Len(Text1.Text))
Clipboard.SetText Mid(Text1.Text, l00C8 + 1, l00CC)
End Sub

Sub cuta_Click ()
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
l00DC = Text1.SelStart
l00E0 = Text1.SelLength
l00E4 = Mid(Text1.Text, 1, l00DC)
l00E8 = Mid(Text1.Text, l00DC + l00E0 + 1, Len(Text1.Text))
Clipboard.SetText Mid(Text1.Text, l00DC + 1, l00E0)
Text1.Text = l00E4 + l00E8
Text1.SelStart = Len(l00E4)
End Sub

Sub Dir1_Change ()
End Sub

Sub Drive1_Change ()
End Sub

Sub File1_Click ()
End Sub

Sub forec_Click ()
colors.Caption = "ForeColor Selection"
colors.Show
End Sub

Sub Form_Load ()
Dim l00F0 As Variant
extsub1050 drawman.hWnd, -1, 0, 0, 0, 0, 3
l00F0 = Chr(13) + Chr(10)
Text1.Text = "__                   _ _   _ _      _ ____      _____   ______" + l00F0
Text1.Text = Text1.Text + "\¯¯ \   _____   /¯¯/ /¯¯¯\     /l¯¯¯¯¯\   /l  ¯¯ ¯ l/l_ _      ll" + l00F0
Text1.Text = Text1.Text + "l \    \/¯  ¯¯¯ \/    / /   /\   \\  l l          l l l   l¯¯¯//__ /    /" + l00F0
Text1.Text = Text1.Text + " \  \        /\        ///     ¯     \ l l    /\    \ l l   l¯¯¯/ ¯/     /_" + l00F0
Text1.Text = Text1.Text + "   \  \___/  \___/ /__/\¯¯ /\__\ l__l  \__\l l   ¯¯¯ l/l______ll" + l00F0
Text1.Text = Text1.Text + "     \/___/\ /__/   \__\/¯¯ \/__//_ /  \/__/l/¯¯ ¯¯¯//__ ___/" + l00F0
Text1.Text = Text1.Text + "      ¯¯¯     ¯¯ ¯  ¯¯        ¯¯   ¯¯   ¯¯   ¯¯¯¯¯¯  ¯¯¯ ¯¯"
Label1.Caption = "30"
HScroll1.Value = 30
End Sub

Sub Form_Resize ()
Panel3D2.Move drawman.Width - 1250
Label1.Move drawman.Width - 1200
Label2.Move drawman.Width - 1200
HScroll1.Move drawman.Width - 1200
On Error Resume Next
Panel3D1.Width = drawman.Width - 1375
Text1.Width = drawman.Width - 1530
On Error Resume Next
Panel3D1.Height = drawman.Height - 1030
Text1.Height = drawman.Height - 1190
Check3D1.Width = drawman.Width - 1600
Check3D1.Top = drawman.Height - 950
End Sub

Sub Form_Unload (p1 As Integer)
Unload macsave
End Sub

Sub HScroll1_Change ()
Label1.Caption = Str(HScroll1)
End Sub

Sub List1_Click ()
End Sub

Sub pastea_Click ()
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
l010E = Text1.SelStart
l0112 = Text1.SelLength
l0116 = Mid(Text1.Text, 1, l010E)
l011A = Mid(Text1.Text, l010E + 1 + l0112, Len(Text1.Text))
Text1.Text = l0116 + Clipboard.GetText() + l011A
Text1.SelStart = Len(l0116) + Len(Clipboard.GetText())
End Sub

Sub Text3_KeyPress (p011E As Integer)
Dim l0120 As Variant
l0120 = Chr(p011E)
End Sub

Sub sub9980 (p0124 As Integer)
Select Case p0124
Case 0
Unload drawman
Case 1
If  drawman.WindowState = 2 Then
drawman.WindowState = 0
Else
drawman.WindowState = 2
End If
Case 2
drawman.WindowState = 1
End Select
End Sub
