VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Esempio 3D in VB5"
   ClientHeight    =   5730
   ClientLeft      =   1095
   ClientTop       =   1770
   ClientWidth     =   8070
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "SUPNAS.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   Begin ComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "Imgl"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "END"
            Object.ToolTipText     =   "Fine"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Apri"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "WIRE"
            Object.ToolTipText     =   "Struttura a Filo"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PITTORE"
            Object.ToolTipText     =   "Resa Sfumata"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DIS"
            Object.ToolTipText     =   "Ridisegna"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pict 
      BackColor       =   &H80000009&
      Height          =   4575
      Left            =   300
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   1
      Top             =   480
      Width           =   7395
   End
   Begin MSComDlg.CommonDialog CMd1 
      Left            =   -180
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      Filter          =   "*.DAT"
   End
   Begin ComctlLib.ImageList Imgl 
      Left            =   -180
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SUPNAS.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SUPNAS.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SUPNAS.frx":0EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SUPNAS.frx":1218
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SUPNAS.frx":1532
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File_m 
      Caption         =   "File"
      Begin VB.Menu Open_i 
         Caption         =   "Apri"
      End
      Begin VB.Menu Exit_i 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu Render_m 
      Caption         =   "Render"
      Begin VB.Menu Wire 
         Caption         =   "Struttura a Filo"
      End
      Begin VB.Menu Painter_i 
         Caption         =   "Render Sfumato"
      End
      Begin VB.Menu Ridis 
         Caption         =   "Ridisegna"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Metodo As Integer ' 0 Render a Filo, 1 = Render Sfufato
Dim FileOpen As String





Sub Render()
 If Len(FileOpen) = 0 Then
    MsgBox "Selezionare un file coordinate Valido!", 16
    Exit Sub
 End If
 
 Pict.Cls
 Pittore Pict, Metodo

End Sub

Private Sub Exit_i_Click()
  End
End Sub

Private Sub Form_Load()
  
  X__max = Pict.ScaleWidth
  Y__max = Pict.ScaleHeight

End Sub






Private Sub Open_i_Click()
 
 CMd1.Filter = "File Coordinate|*.DAT"
 CMd1.ShowOpen
 
 If Len(CMd1.filename) = 0 Then Exit Sub
 
 FileOpen = CMd1.filename
 If Not LoadFile(CMd1.filename) Then
    MsgBox "File Non valido", 16, "Errore Apertura"
    FileOpen = ""
    Exit Sub
 End If
 
 SetColors Form1.Pict
 
 Render
 
End Sub

Private Sub Painter_i_Click()
  Wire.Checked = False
  Painter_i.Checked = True
  Metodo = 1
End Sub

Private Sub Ridis_Click()
  Render
End Sub

Private Sub TB_ButtonClick(ByVal Button As ComctlLib.Button)
   
   Select Case Button.Key
     Case "END"
       Exit_i_Click
     Case "OPEN"
       Open_i_Click
     Case "WIRE"
       Wire_Click
     Case "PITTORE"
       Painter_i_Click
     Case "DIS"
       Render
   End Select
     
End Sub



Private Sub Wire_Click()
  Painter_i.Checked = False
  Wire.Checked = True
  Metodo = 0
End Sub


