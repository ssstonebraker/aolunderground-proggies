' AWAY.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then MsgBox "Whats the Reason?", 16, "Error": Exit Sub
If  Command3D1.Caption = "Start aF|ã" Then
Command3D1.Caption = "Stop"
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Currently aF|ã - Gone Less Than a Minute§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Reason: " + Text1 + "§]v^ñï")
Text1.Enabled = False
Panel3D1.Caption = "0"
Timer1.Enabled = True
Exit Sub
End If
If  Command3D1.Caption = "Stop" Then
Command3D1.Caption = "Start aF|ã"
Text1.Enabled = True
Timer1.Enabled = False
Panel3D1.Caption = "0"
Panel3D3.Caption = "0"
Call subBF0("ïñ^v[§I am Not aF|ã AnyMore§]v^ñï")
End If
End Sub

Sub Command3D2_Click ()
Me.Hide
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Timer1_Timer ()
If  Panel3D3.Caption = "0" Then
Panel3D3.Caption = "60"
Panel3D1.Caption = Panel3D1.Caption + 1
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Currently aF|ã - Gone " + Panel3D1 + " Min(s)§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Reason: " + Text1 + "§]v^ñï")
End If
Panel3D3.Caption = Panel3D3.Caption - 1
If  Text1.Enabled = True Then
Timer1.Enabled = False
Exit Sub
End If
subD08 1
End Sub
