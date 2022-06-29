' FIGHT2.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
l001C = fnA68("«‹- -– Fight Has Been Canceled! –-- -›»")
fightbot.Timer1.Enabled = False
fightbot.Timer1.Interval = 0
Label7.Caption = "Round: 1"
Label8.Caption = "Lost: 0"
Label9.Caption = "Lost: 0"
Label10.Caption = "Won: 0"
Label11.Caption = "Won: 0"
Label2.Caption = "0"
Label4.Caption = "0"
Picture1.Cls
Picture2.Cls
Unload fight2
fightbot.Show
End Sub

Sub Form_Load ()
extsub1050 fight2.hWnd, -1, 0, 0, 0, 0, 3
End Sub
