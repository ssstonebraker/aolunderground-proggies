VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4170
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Loading"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2535
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00000000&
         Caption         =   "Copyright: 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00000000&
         Caption         =   "Company: KnK Founders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Win95/98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5820
         TabIndex        =   3
         Top             =   2700
         Width           =   1035
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5580
         TabIndex        =   4
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Server Helper"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   2760
         TabIndex        =   6
         Top             =   1140
         Width           =   4200
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "KnK 4 Life"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3120
         TabIndex        =   5
         Top             =   705
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'If UserSN() = "" Then
'MsgBox "error:  Please sign on and try again.  Also this program is only for America Online 4.o", vbExclamation, "error"
'End
'End If
End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

End Sub

Private Sub imgLogo_Click()
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

End Sub

Private Sub lblCompany_Click()
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

End Sub
