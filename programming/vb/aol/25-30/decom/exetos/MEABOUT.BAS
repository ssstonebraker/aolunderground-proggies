' MEABOUT.FRM
Option Explicit

Sub Command3D1_Click ()
subD08 .3#
Unload Me
End Sub

Sub Command3D2_Click ()
subD08 .5#
Unload Me
Form72.Show
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Panel3D1_DragDrop (Source As Control, X As Single, Y As Single)
TOSer.mnuespec.Enabled = True
MsgBox "You Can Use the E-Special TOS Method", 77, "Yay!"
End Sub
