VERSION 2.00
Begin Form form10 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPSS Mass Mailer by: EvicTioN"
   ClientHeight    =   4068
   ClientLeft      =   3348
   ClientTop       =   2988
   ClientWidth     =   6840
   ForeColor       =   &H00000000&
   Height          =   4404
   Icon            =   MMER.FRX:0000
   Left            =   3300
   LinkTopic       =   "Form4"
   ScaleHeight     =   4068
   ScaleWidth      =   6840
   Top             =   2700
   Visible         =   0   'False
   Width           =   6936
   Begin TextBox Text4 
      BackColor       =   &H00FFFFFF&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6360
      TabIndex        =   36
      Text            =   "2"
      Top             =   3720
      Width           =   255
   End
   Begin DBPush DBPush15 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Notify"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   360
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:0302
      PictureDown     =   MMER.FRX:126E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   35
      Top             =   2280
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush12 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Room Bust"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   240
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:21DA
      PictureDown     =   MMER.FRX:3146
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   34
      Top             =   2520
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Start when done"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2880
      Top             =   4560
   End
   Begin DBPush DBPush13 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Add Room"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   4440
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:40B2
      PictureDown     =   MMER.FRX:501E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   27
      Top             =   3120
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush11 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Load List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   5640
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:5F8A
      PictureDown     =   MMER.FRX:6EF6
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   26
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush10 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Save List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   4440
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:7E62
      PictureDown     =   MMER.FRX:8DCE
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   25
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush9 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Clear List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   5640
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:9D3A
      PictureDown     =   MMER.FRX:ACA6
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   24
      Top             =   3120
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin PictureBox Picture2 
      Height          =   2725
      Left            =   4440
      Picture         =   MMER.FRX:BC12
      ScaleHeight     =   2690.11
      ScaleMode       =   0  'User
      ScaleWidth      =   2268
      TabIndex        =   20
      Top             =   360
      Width           =   2295
      Begin ListBox List2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
   End
   Begin OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Disabled"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Value           =   -1  'True
      Width           =   1455
   End
   Begin OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Enabled"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin DBPush DBPush8 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Add Name"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   1920
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:1FCEE
      PictureDown     =   MMER.FRX:20C5A
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   16
      Top             =   3120
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin CommonDialog CMDialog1 
      Left            =   840
      Top             =   4560
   End
   Begin DBPush DBPush7 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Save List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   1920
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:21BC6
      PictureDown     =   MMER.FRX:22B32
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   15
      Top             =   3600
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush6 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Load List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   3120
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:23A9E
      PictureDown     =   MMER.FRX:24A0A
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   14
      Top             =   3600
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush3 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Clear List"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   2520
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:25976
      PictureDown     =   MMER.FRX:268E2
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   0
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin PictureBox Picture1 
      Height          =   2725
      Left            =   1920
      Picture         =   MMER.FRX:2784E
      ScaleHeight     =   2690.11
      ScaleMode       =   0  'User
      ScaleWidth      =   2268
      TabIndex        =   13
      Top             =   360
      Width           =   2295
      Begin ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin DBPush DBPush5 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Stop Bot"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   360
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:3B92A
      PictureDown     =   MMER.FRX:3C896
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   12
      Top             =   1800
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush4 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Exit Bot"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   480
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:3D802
      PictureDown     =   MMER.FRX:3E76E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   11
      Top             =   2040
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush2 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Add Room"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   3120
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:3F6DA
      PictureDown     =   MMER.FRX:40646
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   10
      Top             =   3120
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush1 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Start Bot"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   240
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MMER.FRX:415B2
      PictureDown     =   MMER.FRX:4251E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   9
      Top             =   1560
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin TextBox Text2 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Remove Me"
      Top             =   840
      Width           =   1455
   End
   Begin TextBox Text1 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "UPSS MM"
      Top             =   240
      Width           =   1455
   End
   Begin TextBox Text3 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Text            =   "10"
      Top             =   1200
      Width           =   375
   End
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   120
      MessageCount    =   MMER.FRX:4348A
      MessageList     =   MMER.FRX:4348C
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   4320
      Width           =   420
   End
   Begin Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   4320
      Width           =   1095
   End
   Begin Label Label12 
      Height          =   135
      Left            =   1560
      TabIndex        =   32
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin Label Label11 
      Caption         =   "0"
      Height          =   135
      Left            =   1680
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label10 
      Caption         =   "1"
      Height          =   135
      Left            =   1440
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Minutes/Room:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   28
      Top             =   3720
      Width           =   1935
   End
   Begin Label Label6 
      BackColor       =   &H00000000&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Roomz to go to:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
   Begin Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Room Manager"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "PeePz on MM:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Trigger Phrase:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Removal Phrase:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Minutes:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
' MMER.FRM
Option Explicit
Const mc0034 = 9 ' &H9%
Const mc0040 = 2 ' &H2%
Const mc0054 = 1036 ' &H40C%
Const mc0062 = 6 ' &H6%
Const mc0136 = 12 ' &HC%
Const mc013A = 258 ' &H102%
Const mc013E = 16 ' &H10%

Sub DBPush1_Click ()
Dim l001C As Variant
Dim l0028 As Integer
Dim l002C As Integer
Dim l0032 As Integer
' Const mc0034 = 9 ' &H9%
Dim l003A As Integer
Dim l003E As Integer
' Const mc0040 = 2 ' &H2%
Dim l0044 As Integer
Dim l0048 As Variant
Dim l0052 As Integer
' Const mc0054 = 1036 ' &H40C%
Dim l0058 As Variant
' Const mc0062 = 6 ' &H6%
Dim l0064 As Integer
Dim l0066 As Integer
Dim l0068 As Integer
l001C = fn0101()
If  l001C = 0 Then
MsgBox "get into a chat please"
Exit Sub
End If
If  Text3.Text = "" Or Text3.Text = "0" Then
MsgBox "Enter a real time for the bot"
Exit Sub
End If
sub0182
l0028% = extfn04BA("AOL FRAME25", 0&)
l002C% = extfn0510(l0028%, "Welcome, ")
If  l002C% = 0 Then
MsgBox "Not signed on!"
Exit Sub
End If
l0028% = extfn04BA("AOL Frame25", 0&)
If  Form1.Option1.Value = True Then
l0032% = extfn0510(l0028%, "New Mail")
gv00E2 = extfn0538(l0032%, mc0034)
End If
If  Form1.Option4.Value = True Then
l0032% = extfn0510(l0028%, "Old Mail")
gv00E2 = extfn0538(l0032%, mc0034)
End If
If  Form1.Option6.Value = True Then
l0032% = extfn0510(l0028%, "Outgoing Mail")
gv00E2 = extfn0538(l0032%, mc0034)
End If
If  Form1.Option5.Value = True Then
l0032% = extfn0510(l0028%, "Incoming FlashMail")
gv00E2 = extfn0538(l0032%, mc0034)
End If
If  Form1.Option1.Value = True Then
Call sub020F(l0028%, "Read &new Mail")
End If
If  Form1.Option4.Value = True Then
Call sub020F(l0028%, "Check Mail You've &Read")
End If
If  Form1.Option6.Value = True Then
Call sub020F(l0028%, "Check Mail You've &Sent")
End If
If  Form1.Option5.Value = True Then
Call sub020F(l0028%, "Read &Incoming Mail")
End If
Do
DoEvents
l003A% = extfn04BA("#32770", "America Online")
If  l003A% <> 0 Then
l003E% = extfn0524(l003A%, "Static")
l0044% = extfn055D(l003E%, mc0040)
l0048 = fn014A(l0044%)
Label13.Caption = "" + l0048 + ""
If  Label13.Caption = "You have no unread mail." Then
sub01AD
MsgBox "You have no mail!"
Label13.Caption = ""
Exit Sub
End If
End If
l0028% = extfn04BA("AOL FRAME25", 0&)
If  Form1.Option1.Value = True Then
l0032% = extfn0510(l0028%, "New Mail")
End If
If  Form1.Option4.Value = True Then
l0032% = extfn0510(l0028%, "Old Mail")
End If
If  Form1.Option6.Value = True Then
l0032% = extfn0510(l0028%, "Outgoing Mail")
End If
If  Form1.Option5.Value = True Then
l0032% = extfn0510(l0028%, "Incoming FlashMail")
End If
Loop Until l0032% <> 0
sub0274
l0028% = extfn04BA("AOL Frame25", 0&)
Do
DoEvents
If  Form1.Option1.Value = True Then
l0032% = extfn0510(l0028%, "New Mail")
End If
If  Form1.Option4.Value = True Then
l0032% = extfn0510(l0028%, "Old Mail")
End If
If  Form1.Option6.Value = True Then
l0032% = extfn0510(l0028%, "Outgoing Mail")
End If
If  Form1.Option5.Value = True Then
l0032% = extfn0510(l0028%, "Incoming FlashMail")
End If
Loop Until l0032% <> 0
l0052% = extfn0524(l0028%, "_AOL_TREE")
l0058 = extfn02DF(l0052%, mc0054, 0, 0)
If  l0058 < 1 Then
MsgBox "nomail"
Exit Sub
End If
Label11 = "" & l0058 & ""
gv00E2 = extfn0538(l0032%, mc0062)
l0064% = extfn0524(l0028%, "MDIClient")
l0066% = extfn0524(l0064%, "AOL CHILD")
l0068% = extfn0524(l0066%, "_AOL_VIEW")
VBMsg1.SubClasshWnd = l0068%
If  Text3.Text = "" Then
MsgBox "enter a time"
Exit Sub
End If
DBPush1.Enabled = False
DBPush5.Enabled = True
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MMer by: EvicTioN"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• " + Label11 + " Mails going out in " + Text3.Text + " minutes"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text1.Text + Chr(34) + " to be added"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text2.Text + Chr(34) + " to be Removed"
Timer1.Enabled = True
Do
DoEvents
Loop Until Timer1.Enabled = False
If  DBPush5.Enabled = False Then Exit Sub
If  options.Check3.Value = 1 Then
If  List1.ListCount = 0 Then
MsgBox "get some peeps on the mm first"
DBPush1.Enabled = True
DBPush5.Enabled = False
Exit Sub
End If
End If
If  Check1.Value = 1 Then
If  Form1.Option7.Value = True Then
Form3.Show
Me.Hide
Exit Sub
End If
Me.Hide
Load Form2
Exit Sub
End If
End Sub

Sub DBPush10_Click ()
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As String
Dim l00B0 As String
If  List2.ListCount = 0 Then MsgBox "There are no names.", 48: Exit Sub
On Error Resume Next
CMDialog1.Flags = &H4& Or &H2& Or &H800&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Save Names as ..."
CMDialog1.Filter = "Room List (*.RML)|*.RML|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 2
Select Case Err
    Case 0:
    l00A6 = List2.ListCount - 1
    For l00AA = 0 To l00A6 - 1
    l00AE$ = List2.List(l00AA)
    l00B0$ = l00B0$ + l00AE$ & ","
    Next l00AA
    l00B0$ = l00B0$ + List2.List(l00A6)
        Open CMDialog1.Filename For Output As #1
            Print #1, l00B0$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Sub DBPush11_Click ()
Dim l00BE As Variant
Dim l00C2 As String
Dim l00C4 As Integer
On Error Resume Next
CMDialog1.Flags = &H1000& Or &H4& Or &H800& Or &H4000&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Load Name List"
CMDialog1.Filter = "Room List (*.RML)|*.RML|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 1
Select Case Err
    Case 0:
    List1.Clear
    Open CMDialog1.Filename For Input As #1
l00BE = 0
    Do
    Input #1, l00C2$
    l00C2$ = Trim(l00C2$)
    If  l00C2$ <> "" Then List1.AddItem l00C2$
    l00C4% = DoEvents()
l00BE = l00BE + 1
    Loop Until EOF(1)
  Close #1
End Select
Exit Sub
End Sub

Sub DBPush12_Click ()
Form12.Show
End Sub

Sub DBPush13_Click ()
Dim l00CA As Variant
Dim l00CE As Variant
l00CE = InputBox("Room Name Here", "Add what room?")
If  l00CE = "" Then Exit Sub
For l00CA = 0 To List2.ListCount - 1
If  UCase$(l00CE) = UCase$(List2.List(l00CA)) Then Exit Sub
Next l00CA
List2.AddItem l00CE
End Sub

Sub DBPush15_Click ()
If  Timer1.Enabled = True Then
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MMer by: EvicTioN"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• " + Label11 + " Mails going out in " + Text3.Text + " minutes"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text1.Text + Chr(34) + " to be added"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text2.Text + Chr(34) + " to be Removed"
End If
If  Timer1.Enabled = False Then
MsgBox "Bot must be started to use"
End If
End Sub

Sub DBPush2_Click ()
Dim l00D2 As Integer
l00D2% = fn0101()
If  l00D2% = 0 Then
MsgBox "Must be in chatroom to use"
Exit Sub
End If
sub0095 List1
Label6.Caption = List1.ListCount
End Sub

Sub DBPush3_Click ()
Dim l00DA As Variant
l00DA = MsgBox("Are you sure you want to Clear the list?", 4, "Clear MM list")
If  l00DA = 6 Then
List1.Clear
Exit Sub
End If
If  l00DA = 7 Then Exit Sub
End Sub

Sub DBPush4_Click ()
Dim l00DE As Variant
Timer1.Enabled = False
VBMsg1.SubClasshWnd = 0
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MassMailer Bot Is Closed!!"
DBPush1.Enabled = True
DBPush5.Enabled = False
Form1.List1.Clear
For l00DE = 0 To List1.ListCount - 1
Form1.List1.AddItem List1.List(l00DE)
Next l00DE
List1.Clear
Form1.Label6.Caption = Form1.List1.ListCount
Form1.WindowState = 0
Unload Me
End Sub

Sub DBPush5_Click ()
Timer1.Enabled = False
VBMsg1.SubClasshWnd = 0
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MassMailer Bot Is Disabled!!"
DBPush1.Enabled = True
DBPush5.Enabled = False
End Sub

Sub DBPush6_Click ()
Dim l00E2 As Variant
Dim l00E6 As String
Dim l00E8 As Integer
On Error Resume Next
CMDialog1.Flags = &H1000& Or &H4& Or &H800& Or &H4000&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Load Name List"
CMDialog1.Filter = "MM List (*.MML)|*.MML|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 1
Select Case Err
    Case 0:
    List1.Clear
    Open CMDialog1.Filename For Input As #1
l00E2 = 0
    Do
    Input #1, l00E6$
    l00E6$ = Trim(l00E6$)
    If  l00E6$ <> "" Then List1.AddItem l00E6$
    l00E8% = DoEvents()
l00E2 = l00E2 + 1
    Loop Until EOF(1)
  Close #1
End Select
Label6.Caption = List1.ListCount
Exit Sub
End Sub

Sub DBPush7_Click ()
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As String
Dim l00F4 As String
If  List1.ListCount = 0 Then MsgBox "There are no names.", 48: Exit Sub
On Error Resume Next
CMDialog1.Flags = &H4& Or &H2& Or &H800&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Save Names as ..."
CMDialog1.Filter = "MM List (*.MML)|*.MML|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 2
Select Case Err
    Case 0:
    l00EA = List1.ListCount - 1
    For l00EE = 0 To l00EA - 1
    l00F2$ = List1.List(l00EE)
    l00F4$ = l00F4$ + l00F2$ & ","
    Next l00EE
    l00F4$ = l00F4$ + List1.List(l00EA)
        Open CMDialog1.Filename For Output As #1
            Print #1, l00F4$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Sub DBPush8_Click ()
Dim l00F6 As Variant
Dim l00FA As Variant
l00FA = InputBox("Screen Name Here", "Add who?")
If  l00FA = "" Then Exit Sub
For l00F6 = 0 To List1.ListCount - 1
If  UCase$(l00FA) = UCase$(List1.List(l00F6)) Then Exit Sub
Next l00F6
List1.AddItem l00FA
Label6.Caption = List1.ListCount
End Sub

Sub DBPush9_Click ()
Dim l00FE As Variant
l00FE = MsgBox("Are you sure you want to Clear the list?", 4, "Clear MM list")
If  l00FE = 6 Then
List1.Clear
Text3.Text = (Text4.Text * List2.ListCount) + Text4.Text
Exit Sub
End If
If  l00FE = 7 Then Exit Sub
End Sub

Sub Form_Activate ()
sub0193 Me
Label6.Caption = List1.ListCount
End Sub

Sub Form_Load ()
sub0193 Me
form10.Width = 4300
List2.AddItem "poa"
List2.AddItem "poa2"
List2.AddItem "poa3"
List2.AddItem "fate"
List2.AddItem "fate2"
List2.AddItem "fate3"
List2.AddItem "server"
List2.AddItem "server2"
List2.AddItem "server3"
List2.AddItem "server4"
List2.AddItem "server5"
List2.AddItem "server6"
List2.AddItem "server7"
List2.AddItem "server8"
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
Label6.Caption = List1.ListCount
Label6 = Form1.Label6
End Sub

Sub List2_DblClick ()
If  List2.ListCount = 1 Then
MsgBox "You have to have at least 1 room!"
Exit Sub
End If
List2.RemoveItem List2.ListIndex
Text3.Text = (Text4.Text * List2.ListCount) + Text4.Text
End Sub

Sub Option1_Click ()
form10.Width = 6980
Text3.Text = (Text4.Text * List2.ListCount) + Text4.Text
End Sub

Sub Option2_Click ()
form10.Width = 4300
End Sub

Sub Text3_Change ()
If  Text3.Text < "0" Then
Text3.Text = "1"
MsgBox "must be set above 0"
End If
End Sub

Sub Text4_Change ()
If  Text4.Text <> "" Then
Text3.Text = (Text4.Text * List2.ListCount) + Text4.Text
End If
If  Text4.Text < 0 Or Text4.Text = 0 Then
Text4.Text = "1"
MsgBox "must be set above 0"
End If
End Sub

Sub Timer1_Timer ()
Dim l010C As Variant
Dim l0110 As String
Dim l0114 As Integer
Dim l0116 As Integer
Dim l0118 As Variant
Dim l011C As Integer
Dim l011E As Variant
Dim l0124 As Variant
Dim l0128 As Integer
Dim l0130 As Variant
' Const mc0136 = 12 ' &HC%
' Const mc013A = 258 ' &H102%
' Const mc013E = 16 ' &H10%
Dim l0140 As Integer
l0128% = extfn04BA("AOL FRAME25", 0&)
If  Option1.Value = True Then
    If  List2.ListCount <> 0 Then
            For l010C = 0 To List2.ListCount - 1
            List2.Selected(l010C) = True
            Next l010C
            l0130 = List2.Text
    sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MMer Going To PR: " & l0130 & ""
    l0124 = 1
    l0110$ = fn00CD()
L8C1E:
    If  l011E = 1 Then Exit Sub
        Call sub020F(l0128%, "Keyword...")
            Do
            DoEvents
            If  l011E = 1 Then Exit Sub
            l0114% = extfn0510(l0128%, "keyword")
            Loop Until l0114% > 0
        l0116% = extfn0524(l0114%, "_AOL_Edit")
        l0118 = extfn02F3(l0116%, mc0136, 0, "aol://2719:2-2-" & l0130)
        l0118 = extfn0398(l0116%, mc013A, 13, 0)
            Do
            DoEvents
            If  l011E = 1 Then Exit Sub
            l011C% = extfn04BA("#32770", "America Online")
            Loop Until l011C% > 0 Or fn00CD() <> l0110$
                If  l011C% > 0 Then
                If  l011E = 1 Then Exit Sub
                l011C% = extfn04BA("#32770", "America Online")
                l0118 = extfn0398(l011C%, mc013E, 0, 0)
                GoTo L8C1E
                ElseIf fn00CD() <> l0110$ Then
                If  l011E = 1 Then Exit Sub
                List2.RemoveItem List2.ListIndex
                l0124 = 0
                l0140% = Text3.Text - Text4.Text
                sub0269 (.8#)
                GoTo L8F56
                End If
        End If
End If
L8F56:
Text3 = Text3 - 1
If  Text3.Text = "0" Then
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MMer is now Mailing!"
VBMsg1.SubClasshWnd = 0
Timer1.Enabled = False
Exit Sub
End If
sub0263 "²º·‚¸‚•²º·‚¸‚• UPSS MMer by: EvicTioN"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• " + Label11 + " Mails going out in " + Text3.Text + " minutes"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text1.Text + Chr(34) + " to be added"
sub0269 (.001#)
sub0263 "²º·‚¸‚•²º·‚¸‚• Say " + Chr(34) + Text2.Text + Chr(34) + " to be Removed"
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim l014E As Variant
Dim l0154 As String
Dim l0156 As String
Dim l0158 As String
Dim l015C As Variant
Dim l0160 As Variant
l0154$ = extfn04A2$(lParam)
l0156$ = Mid$(l0154$, 3, InStr(l0154$, ":") - 3)
l0158$ = Mid$(l0154$, InStr(l0154$, ":") + 2)
If  UCase(l0158$) = UCase(Text1.Text) Then
If  l0156$ = fn0141() Then Exit Sub
For l014E = 0 To List1.ListCount - 1
If  UCase$(l0156$) = UCase$(List1.List(l014E)) Then Exit Sub
If  l0156$ = fn0141() Then Exit Sub
Next l014E
List1.AddItem l0156$
Call sub0269(1)
sub0263 "²º·‚¸‚• " & l0156$ & " is added to the Mass Mail."
Label6.Caption = List1.ListCount
End If
If  UCase(l0158$) = UCase(Text2.Text) Then
If  l0156$ = fn0141() Then Exit Sub
For l015C = 0 To List1.ListCount - 1
l0160 = InStr(l0156$, List1.List(l015C))
List1.RemoveItem l015C
GoTo L9440
Next l015C
L9440:
Call sub0269(1)
sub0263 "²º·‚¸‚• " & l0156$ & " is Off the Mass Mail."
Label6.Caption = List1.ListCount
End If
End Sub
