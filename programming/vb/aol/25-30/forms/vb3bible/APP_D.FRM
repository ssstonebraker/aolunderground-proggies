VERSION 2.00
Begin Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   930
   ClientTop       =   1575
   ClientWidth     =   6465
   Height          =   6525
   Icon            =   APP_D.FRX:0000
   Left            =   855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   6465
   Top             =   1215
   Width           =   6615
End
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$) As Integer

Const APPNAME = "Waite Group - Appendix D"
Dim MyColor As Integer

Sub Form_Click ()

     MyColor = (MyColor + 1) Mod 15
     Me.BackColor = QBColor(MyColor)

End Sub

Sub Form_Load ()

     Dim temp As String * 16
     Dim junk As Integer
     
     temp = String$(16, 0)
     junk = GetProfileString(APPNAME, "Left", "1000", temp, 16)
     Me.Left = Val(temp)
     
     temp = String$(16, 0)
     junk = GetProfileString(APPNAME, "Top", "1000", temp, 16)
     Me.Top = Val(temp)

     temp = String$(16, 0)
     junk = GetProfileString(APPNAME, "Width", "5000", temp, 16)
     Me.Width = Val(temp)

     temp = String$(16, 0)
     junk = GetProfileString(APPNAME, "Height", "3000", temp, 16)
     Me.Height = Val(temp)

     Me.BackColor = QBColor(GetProfileInt(APPNAME, "Color", 1))
     
End Sub

Sub Form_Unload (Cancel As Integer)

     Dim junk As Integer

     junk = WriteProfileString(APPNAME, "Left", Str$(Me.Left))
     junk = WriteProfileString(APPNAME, "Top", Str$(Me.Top))
     junk = WriteProfileString(APPNAME, "Width", Str$(Me.Width))
     junk = WriteProfileString(APPNAME, "Height", Str$(Me.Height))
     junk = WriteProfileString(APPNAME, "Color", Str$(MyColor))

End Sub

