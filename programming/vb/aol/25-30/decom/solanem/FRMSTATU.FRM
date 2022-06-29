VERSION 2.00
Begin Form frmStatus 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Status"
   ClientHeight    =   2190
   ClientLeft      =   2955
   ClientTop       =   4260
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2595
   Left            =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4095
   Top             =   3915
   Width           =   4215
   Begin SSCommand cmdStop 
      BevelWidth      =   1
      Caption         =   "Stop"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   645
      Left            =   2895
      TabIndex        =   7
      Top             =   735
      Width           =   795
   End
   Begin SSCommand cmdPause 
      BevelWidth      =   1
      Caption         =   "Pause"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   645
      Left            =   2115
      TabIndex        =   6
      Top             =   735
      Width           =   810
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00000000&
      BevelInner      =   1  'Inset
      Caption         =   "Panel3D1"
      FloodType       =   1  'Left To Right
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   2040
      TabIndex        =   5
      Top             =   660
      Width           =   1725
   End
   Begin SSFrame Frame3D3 
      Caption         =   "Commands:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   1215
      Left            =   1920
      TabIndex        =   4
      Top             =   330
      Width           =   1950
   End
   Begin ListBox lstRemoved 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   -1  'True
      FontUnderline   =   0   'False
      Height          =   450
      Left            =   300
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin SSFrame Frame3D2 
      Caption         =   "Removed:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   1215
      Left            =   195
      TabIndex        =   2
      Top             =   330
      Width           =   1665
   End
   Begin SSPanel pnlPercent 
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      Caption         =   "Panel3D1"
      FloodColor      =   &H0000FF00&
      FloodType       =   1  'Left To Right
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   1635
      Width           =   3735
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Mailing Status"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   1995
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   3975
   End
End

Sub cmdPause_Click ()
If cmdPause.Caption = "Pause" Then
    cmdPause.Caption = "Resume"
Else
    cmdPause.Caption = "Pause"
End If
End Sub

Sub cmdStop_Click ()
cmdStop.Caption = "Ending..."
End Sub

Sub Form_GotFocus ()
stayontop Me
End Sub

Sub Form_Load ()
lstRemoved.Height = 870
stayontop Me
frm% = Me.hWnd: j = setfocusapi(frm%): DoEvents
End Sub

Sub Form_LostFocus ()
stayontop Me
End Sub

Sub Form_Resize ()
stayontop Me
End Sub

Sub lstRemoved_DblClick ()
frmMMer.lstPeople.AddItem lstRemoved.text
lstRemoved.RemoveItem lstRemoved.listindex
End Sub

