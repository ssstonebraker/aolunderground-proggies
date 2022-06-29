' PWAIT.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001E As Variant
Dim l002A As Variant
Dim l0030 As Variant
Dim l0036 As Variant
Dim l003C As Variant
Dim l0044 As Variant
Dim l0048 As String
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As Variant
Dim l005C As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0076 As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A2 As Variant
If  maillist.Option3D1.Value = True Then l001E = "New Mail"
If  maillist.Option3D2.Value = True Then l001E = "Old Mail"
If  maillist.Option3D3.Value = True Then l001E = "Outgoing Mail"
If  maillist.Option3D4.Value = True Then l001E = "Incoming FlashMail"
Panel3D2.Visible = False
l002A = extfn1018("AOL FRAME25", 0&)
l0030 = extfn10F8(l002A, l001E)
l0036 = extfn10C0(l0030, "_AOL_TREE")
l003C = extfnFE0(l0036, 1036, 0, 0)
pwait.Label1.Caption = "Collecting Mailbox..."
DoEvents
For l0044 = 0 To l003C - 1
DoEvents
l0048$ = String(255, 0)
l004C = extfn11A0(l0036, 1034, l0044, l0048$)
l0050 = InStr(l0048$, Chr(9))
l0054 = Mid(l0048$, l0050 + 1, Len(l0048$))
l0058 = InStr(l0054, Chr(9))
l0048$ = Mid(l0054, l0058 + 1, Len(l0054))
l005C = Chr(13) + Chr(10)
Text1 = Text1 + l005C + Str(l0044) + ".) " + l0048$
Next l0044
l0068 = fn5D0("&Mail", "&Compose Mail")
40 :
DoEvents
l006C = extfn10F8(l002A, "Compose Mail")
If  l006C = 0 Then GoTo 40
pwait.Label1.Caption = "Setting information..."
l0070 = extfn10F8(l006C, "To:")
l0076 = extfn1168(l0070, 2)
l0080 = extfn11A0(l0076, 12, 0, Text2.Text)
l0084 = extfn10F8(l006C, "Subject:")
l0088 = extfn1168(l0084, 2)
l008C = extfn11A0(l0088, 12, 0, "«‹- -– Fate Ultra 9.0 Mail Box List - " + l001E + " –-- -›»")
l0090 = extfn1168(l0070, 1)
l0094 = extfn11A0(l0090, 12, 0, Text1.Text)
DoEvents
l0098 = extfn10F8(l006C, "Send")
l009C = extfn1168(l0098, 2)
l00A2 = fn1480(l009C)
Unload pwait
End Sub

Sub Form_Load ()
extsub1050 pwait.hWnd, -1, 0, 0, 0, 0, 3
End Sub
