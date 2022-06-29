VERSION 4.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Words from PoOH"
   ClientHeight    =   105
   ClientLeft      =   1665
   ClientTop       =   90
   ClientWidth     =   6690
   Height          =   510
   Icon            =   "Form5.frx":0000
   Left            =   1605
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":030A
   ScaleHeight     =   105
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   -255
   Width           =   6810
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Activate()
Do
Form5.Height = Form5.Height + 5
Loop Until Form5.Height > 6545
Dim mess As String
Dim a As Integer
mess = mess & "" & Chr(13)
mess = mess & "      OreO¹·°   ƒòr ÀöL 4·º " & Chr(13)
mess = mess & "By: PôOH and FuSH Yu MaNG" & Chr(13)
mess = mess & Chr(13)
mess = mess & "Sup Ya'll, Itz me the Old Pogi G. I changed my name because peeps kept on terming me. So my new handle is PoOH." & Chr(13)
mess = mess & "I dont program for AOL 3.0 no more, because it got old, so now I am makin shit for AOL 4.0 . " & Chr(13)
mess = mess & "We might make OreO V1.1, Depending on if I got time. Hopefully me and Sky will get back together and make a AOL 4.0 Version Of Optik." & Chr(13)
mess = mess & "WEll until this day, I still got like 5 or 6 progs I didnt finish, so look out for them" & Chr(13)
mess = mess & "Well thats most of the INFO for now ya'll" & Chr(13)
mess = mess & "                   Lata" & Chr(13)
Playwav ("lostones")
For a = 1 To Len(mess)
Label1 = Mid$(mess, 1, a)
Call timeout(0.1)
Next a
If Len(Label1) = Len(mess) Then
Unload Form5
Form2.Show
End If
End Sub

Private Sub Form_Load()
Form2.Hide
StayOnTop Me
Form5.Top = 0
Label1 = ""
End Sub

