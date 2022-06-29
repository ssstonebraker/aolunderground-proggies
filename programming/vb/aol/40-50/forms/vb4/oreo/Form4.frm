VERSION 4.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   45
   ClientLeft      =   1755
   ClientTop       =   75
   ClientWidth     =   6690
   Height          =   450
   Icon            =   "Form4.frx":0000
   Left            =   1695
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":030A
   ScaleHeight     =   45
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   -270
   Width           =   6810
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   6135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Activate()
Do
Form4.Height = Form4.Height + 5
Loop Until Form4.Height > 6960
Dim mess As String
Dim a As Integer
mess = mess & "" & Chr(13)
mess = mess & "      OreO¹·°   ƒòr ÀöL 4·º " & Chr(13)
mess = mess & "By: PôOH and FuSH Yu MaNG" & Chr(13)
mess = mess & Chr(13)
mess = mess & "Sup Ya'll, Itz me the Old Pogi G. I changed my name because peeps kept on terming me. So my new handle is PoOH." & Chr(13)
mess = mess & "I dont program for AOL 3.0 no more, because it got old, so now I am makin shit for AOL 4.0 . " & Chr(13)
mess = mess & "We made OreO with Visual Basic 4.0 Standard Version. I used alot of different OCXs and some of are other shit." & Chr(13)
mess = mess & "We might make OreO V1.1, Depending on if I got time. Hopefully me and Sky will get back together and make a AOL 4.0 Version Of Optik." & Chr(13)
mess = mess & "OH I forgot to tell ya'll that FuSH can make some phat ass art, so if you catch him online, ask him for some art." & Chr(13)
mess = mess & "Well thats most of the INFO for now ya'll" & Chr(13)
mess = mess & "                   Lata" & Chr(13)
Playwav ("lostones")
For a = 1 To Len(mess)
Label1 = Mid$(mess, 1, a)
Call timeout(0.1)
Next a
If Len(Label1) = Len(mess) Then
Unload Form4
Form2.Show
End If
End Sub

Private Sub Form_Load()
Form2.Hide
StayOnTop Me
Form4.Top = 0
Label1 = ""
End Sub


