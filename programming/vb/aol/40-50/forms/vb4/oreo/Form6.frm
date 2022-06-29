VERSION 4.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Words from FuSH Yu MaNG"
   ClientHeight    =   45
   ClientLeft      =   1530
   ClientTop       =   150
   ClientWidth     =   6690
   Height          =   450
   Icon            =   "Form6.frx":0000
   Left            =   1470
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form6.frx":030A
   ScaleHeight     =   45
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   -195
   Width           =   6810
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Activate()
Do
Form6.Height = Form6.Height + 5
Loop Until Form6.Height > 5325
Dim mess As String
Dim a As Integer
mess = mess & "" & Chr(13)
mess = mess & "      OreO¹·°   ƒòr ÀöL 4·º " & Chr(13)
mess = mess & "By: PôOH and FuSH Yu MaNG" & Chr(13)
mess = mess & Chr(13)
mess = mess & "Sup Ya'll, Itz me the FuSh. I dont do that much coding, but I make really good art. PoOH Is trying to help me code so I can Make my Own Big Prog. He's a Real cool guy, so when he is online, Talk to him." & Chr(13)
mess = mess & "If you need art, then Just mail use through this prog. I love That OreO Cookie that PoOh Made for the Intro. " & Chr(13)
mess = mess & "I plan to make a few progs so look out for them." & Chr(13)
mess = mess & "Well thats most of the INFO for now ya'll" & Chr(13)
mess = mess & "                   Lata" & Chr(13)
Playwav ("lostones")
For a = 1 To Len(mess)
Label1 = Mid$(mess, 1, a)
Call timeout(0.1)
Next a
If Len(Label1) = Len(mess) Then
Unload Form6
Form2.Show
End If
End Sub


Private Sub Form_Load()
Form2.Hide
StayOnTop Me
Form6.Top = 0
Label1 = ""
End Sub


