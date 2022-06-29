VERSION 5.00
Begin VB.Form frmLISTFILES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Required Images for Fwogger.exe"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frmLISTFILES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "List of required files"
      Height          =   4770
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4395
      Begin VB.PictureBox gifSCHOOLBUS 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   450
         Picture         =   "frmLISTFILES.frx":0442
         ScaleHeight     =   450
         ScaleWidth      =   1125
         TabIndex        =   17
         Top             =   1665
         Width           =   1185
      End
      Begin VB.CommandButton cmdDOWNLOADgifjpg 
         Caption         =   "Download these six images from WWW"
         Height          =   375
         Left            =   540
         TabIndex        =   16
         Top             =   4185
         Width           =   3120
      End
      Begin VB.CommandButton cmdSAVEICOs 
         Caption         =   "Save these two images"
         Height          =   375
         Left            =   2070
         TabIndex        =   15
         Top             =   675
         Width           =   1860
      End
      Begin VB.PictureBox jpgWATER 
         Height          =   690
         Left            =   2205
         Picture         =   "frmLISTFILES.frx":08BC
         ScaleHeight     =   630
         ScaleWidth      =   1830
         TabIndex        =   14
         Top             =   3375
         Width           =   1890
      End
      Begin VB.PictureBox gifSPORTSCAR 
         AutoSize        =   -1  'True
         Height          =   480
         Left            =   2385
         Picture         =   "frmLISTFILES.frx":2BCC
         ScaleHeight     =   420
         ScaleWidth      =   1440
         TabIndex        =   13
         Top             =   1710
         Width           =   1500
      End
      Begin VB.PictureBox gifLIMMO 
         AutoSize        =   -1  'True
         Height          =   465
         Left            =   135
         Picture         =   "frmLISTFILES.frx":350A
         ScaleHeight     =   405
         ScaleWidth      =   1800
         TabIndex        =   12
         Top             =   2565
         Width           =   1860
      End
      Begin VB.PictureBox gifLOG 
         AutoSize        =   -1  'True
         Height          =   480
         Left            =   2295
         Picture         =   "frmLISTFILES.frx":3D3B
         ScaleHeight     =   420
         ScaleWidth      =   1635
         TabIndex        =   11
         Top             =   2565
         Width           =   1695
      End
      Begin VB.PictureBox jpgGRASS 
         Height          =   690
         Left            =   135
         Picture         =   "frmLISTFILES.frx":4777
         ScaleHeight     =   630
         ScaleWidth      =   1830
         TabIndex        =   10
         Top             =   3375
         Width           =   1890
      End
      Begin VB.PictureBox icoFROGG 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1215
         Picture         =   "frmLISTFILES.frx":5023
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   585
         Width           =   540
      End
      Begin VB.PictureBox icoDEAD 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   225
         Picture         =   "frmLISTFILES.frx":58ED
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   585
         Width           =   540
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "grass.jpg"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   19
         Top             =   3150
         Width           =   630
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "schoolbus.gif"
         Height          =   195
         Index           =   5
         Left            =   540
         TabIndex        =   18
         Top             =   1440
         Width           =   930
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   90
         X2              =   4320
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   90
         X2              =   4320
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "water.jpg"
         Height          =   195
         Index           =   7
         Left            =   2700
         TabIndex        =   7
         Top             =   3150
         Width           =   645
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "sportscar.gif"
         Height          =   195
         Index           =   6
         Left            =   2655
         TabIndex        =   6
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "log.gif"
         Height          =   195
         Index           =   4
         Left            =   2835
         TabIndex        =   5
         Top             =   2340
         Width           =   420
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "limmo.gif"
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   4
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "grass.jpg"
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   3
         Top             =   3645
         Width           =   630
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "frogg.ico"
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblLISTEDFILE 
         AutoSize        =   -1  'True
         Caption         =   "dead.ico"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLISTFILES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long  'Used for URL Coding

Private Sub cmdDOWNLOADgifjpg_Click()
ShellExecute 0, "Open", "http://come.to/magikcube", "", "", 0
End Sub

Private Sub cmdSAVEICOs_Click()
Dim gotoweb As VbMsgBoxResult
Call SavePicture(icoDEAD, App.Path & "\dead.ico")
Call SavePicture(icoFROGG.Picture, App.Path & "\frogg.ico")
gotoweb = MsgBox(App.Path & "\dead.ico" & vbCrLf & App.Path & "\frogg.ico" & vbCrLf & vbCrLf & "Two images have been saved. Would you like to access the internet to download the other 6 images?", vbYesNo + vbQuestion, "Saved Pictures...")
If gotoweb = vbYes Then cmdDOWNLOADgifjpg_Click
If gotoweb = vbNo Then frmLISTFILES.Hide
End Sub
