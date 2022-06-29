' SMARTBOT.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l0036 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
l0036 = fnA30("«‹- -– Fate Ultra - Greeter Bot Enabled –-- -›»")
DoEvents

30 :
l0036 = InStr(1, hello.Text, "%S", 1)
If  l0036 Then
If  l0036 = 1 Then
Text4.Text = ""
Text5.Text = Mid(hello.Text, l0036 + 2, Len(hello.Text))
Else
l004C = Mid(hello.Text, 1, l0036 - 1)
Text4.Text = l004C
l0050 = Mid(hello.Text, l0036 + 2, Len(hello.Text))
Text5.Text = l0050
End If
Else
If  hello.Text = "" Then
l0054 = MsgBox("You must enter in something, do you wish to use the defaults?", 4 + 32)
Else
l0054 = MsgBox("You must enter in something that has ""%s"", do you wish to use the defaults?", 4 + 32)
End If
If  l0054 = 6 Then
hello.Text = "«‹- -– Fate Ultra - %s Has Entered the Room –-- -›»"
GoTo 30
Else
Exit Sub
End If
End If
Timer1.Enabled = True
Timer1.Interval = 1
Command3D1.Caption = "Stop Bot"
Else
l0036 = fnA30("«‹- -– Fate Ultra - Greeter Bot Disabled –-- -›»")
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Start Bot"
End If
End Sub

Sub Form_Load ()
Dim l0062 As Variant
Dim l006A As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
extsub1050 smartbot.hWnd, -1, 0, 0, 0, 0, 3
l0062 = extfn1018("AOL FRAME25", 0&)
l0062 = extfn10F8(l0062, "AOL Live!")
l0062 = extfn1088(l0062)
l006A = extfn10C0(l0062, "_AOL_LISTBOX")
If  l006A = 0 Then
MsgBox "You must be in a chat room!", 16
Unload smartbot
Exit Sub
End If
l0070 = extfn1168(l006A, 2)
l0074 = extfn1168(l0070, 2)
l007A = fn1480(l0074)
34 :
DoEvents
l007E = extfn1018("_AOL_MODAL", "Chat Preferences")
If  l007E = 0 Then GoTo 34
l0082 = extfn10F8(l007E, "Notify me when")
l0086 = extfn1168(l0082, 2)
l008C = extfnFE0(l0086, 1025, 1, 0)
l0090 = extfn1168(l0086, 3)
l0094 = extfnFE0(l0090, 1025, 1, 0)
l0098 = extfn10F8(l007E, "OK")
l009C = fn1480(l0098)
End Sub

Sub Form_Unload (p1 As Integer)
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As Variant
DoEvents
l00A2 = extfn1018("AOL FRAME25", 0&)
l00A2 = extfn10F8(l00A2, "AOL Live!")
If  l00A2 = 0 Then GoTo 67
l00A2 = extfn1088(l00A2)
l00A6 = extfn10C0(l00A2, "_AOL_LISTBOX")
l00AA = extfn1168(l00A6, 2)
l00AE = extfn1168(l00AA, 2)
l00B2 = fn1480(l00AE)
346 :
DoEvents
l00B6 = extfn1018("_AOL_MODAL", "Chat Preferences")
If  l00B6 = 0 Then GoTo 346
l00BA = extfn10F8(l00B6, "Notify me when")
l00BE = extfn1168(l00BA, 2)
l00C2 = extfnFE0(l00BE, 1025, 0, 0)
l00C6 = extfn1168(l00BE, 3)
l00CA = extfnFE0(l00C6, 1025, 0, 0)
l00CE = extfn10F8(l00B6, "OK")
l00D2 = fn1480(l00CE)
67 :
End Sub

Sub Timer1_Timer ()
Dim l00E0 As Variant
Text1.Text = fn170()
l00E0 = fnB48()
End Sub
