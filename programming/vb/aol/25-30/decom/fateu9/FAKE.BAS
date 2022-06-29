' FAKE.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l0046 As Variant
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005C As Variant
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As String
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "You haven't filled out all the forms!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You haven't filled out all the forms!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "You haven't filled out all the forms!", 16
Exit Sub
End If
If  Text4.Text = "" Then
MsgBox "You haven't filled out all the forms!", 16
Exit Sub
End If
l0046 = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l004E = fn5D0("Mem&bers", "Set Preferences")
Else
l004E = fn5D0("Mem&bers", "Preferences")
End If

15 :
DoEvents
l0052 = extfn10F8(l0046, "Preferences")
If  l0052 = 0 Then GoTo 15
l0056 = extfn10F8(l0052, "Mail")
l005C = extfn1168(l0056, 2)
l0062 = extfnFE0(l005C, &H201, 0, 0)
l0066 = extfnFE0(l005C, &H202, 0, 0)
25 :
DoEvents
l006A = extfn1018("_AOL_MODAL", 0&)
If  l006A = 0 Then GoTo 25
l006E = extfnFE0(l0052, 16, 0, 0)
l0072 = extfn10F8(l006A, "Confirm mail")
l0076 = extfn1168(l0072, 2)
l007A = extfnFE0(l0076, 1025, 0, 0&)
l007E = extfn10F8(l006A, "Close mail")
l0082 = extfn1168(l007E, 3)
l0086 = extfnFE0(l0082, 1025, 1, 0&)
l008A = extfn10F8(l006A, "OK")
l008E = extfnFE0(l008A, &H201, 0, 0)
l0092 = extfnFE0(l008A, &H202, 0, 0)
l0098 = extfn1130(l0046)
l0046 = extfn1018("AOL FRAME25", 0&)
l009C = extfn10F8(l0046, "Welcome, ")
l00A0$ = String(250, 0)
l00A2 = extfnFA8(l009C, l00A0$, 250)
l00A6 = Mid(l00A0$, 10, 10)
l00AA = InStr(l00A6, "!")
If  l00AA Then
l00A6 = Mid(l00A6, 1, l00AA - 1)
End If
l004E = extfn10C0(l0046, "AOL TOOLBAR")
l0052 = extfn10C0(l004E, "_AOL_ICON")
l0056 = extfn1168(l0052, 2)
l005C = fn1480(l0056)
DoEvents
l00B2 = Time$
l00B6 = Date$

10 :
DoEvents
l0052 = extfn10F8(l0046, "Compose Mail")
If  l0052 = 0 Then GoTo 10
l00BA = Date
l00BE = InStr(l00BA, "/")
l00C2 = Mid(l00BA, l00BE + 1, Len(l00BA))
l00C6 = InStr(l00C2, "/")
l00CA = Mid(l00C2, l00C6 + 1, Len(l00C2))
l00CE = Mid(l00BA, 1, l00BE - 1)
l00D2 = Mid(l00C2, 1, l00C6 - 1)
If  Len(l00D2) = 1 Then
l00D2 = "0" + Trim(Str(l00D2))
End If
l00D6 = l00CA + "-" + l00CE + "-" + l00D2
l0056 = extfn10F8(l0052, "To:")
l005C = extfn1168(l0056, 2)
l0062 = extfn11A0(l005C, 12, 0, Text1.Text)
l0066 = extfn10F8(l0052, "Subject:")
l006A = extfn1168(l0066, 2)
l006E = extfn11A0(l006A, 12, 0, "Fwd: " + Text3)
l0072 = extfn1168(l0056, 1)
l00DC = Chr(13) + Chr(10)
l0076 = extfn11A0(l0072, 12, 0, "-----------------" + l00DC + "Forwarded Message:" + l00DC + "Subj:" + Chr(9) + Text3 + l00DC + "Date:" + Chr(9) + l00D6 + " " + l00B2 + " EDT" + l00DC + "From:" + Chr(9) + Text2 + l00DC + "To:" + Chr(9) + l00A6 + l00DC + l00DC + Text4)
l007A = extfn10F8(l0052, "Send")
l007E = extfn1168(l007A, 2)
l0082 = extfnFE0(l007E, &H201, 0, 0)
l0086 = extfnFE0(l007E, &H202, 0, 0)
34 :
DoEvents
If  fn138() = 5 Then
l008A = extfn1018("#32770", "America Online")
Else
l008A = extfn1018("_AOL_MODAL", 0&)
End If
If  l008A = 0 Then GoTo 34
l008E = extfn10F8(l008A, "OK")
l0092 = extfn1130(l008E)
l0098 = extfnFE0(l0092, &H201, 0, 0)
l00E0 = extfnFE0(l0092, &H202, 0, 0)
l00E4 = extfnFE0(l0052, 16, 0, 0)
End Sub

Sub Command3D2_Click ()
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F0 As String
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
l00E8 = extfn1018("AOL FRAME25", 0&)
l00EC = extfn10F8(l00E8, "Welcome")
l00F0$ = String(30, 0)
l00F2 = extfn11A0(l00EC, 13, 255, l00F0$)
l00F6 = Mid(l00F0$, 10, 10)
l00FA = InStr(l00F6, "!")
If  l00FA Then
l00F6 = Mid(l00F6, 1, l00FA - 1)
End If
Text1.Text = "TOSEmail2"
Text2.Text = "Snoop"
Text3.Text = "ATTENTION: Your Account"
l00FE = Chr(13) + Chr(10)
Text4.Text = "Dear User," + l00FE + "    Due to internal error #1785, we have lost vital information concerning on your master account log-on password. Please click on 'Reply' and enter in your password. Thank you for your time and concideration." + l00FE + l00FE + "Sam Kieth - Head of America Online Services"
End Sub

Sub Form_Load ()
extsub1050 fake.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub sub9980 ()
Unload fake
End Sub
