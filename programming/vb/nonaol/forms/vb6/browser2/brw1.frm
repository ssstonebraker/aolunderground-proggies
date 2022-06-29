VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Internet Browser"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "brw1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   1535
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Back"
            Key             =   "back"
            Object.ToolTipText     =   "Back To Last Site"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Home"
            Key             =   "Home"
            Object.ToolTipText     =   "Go to Home Page"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh This Page"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop Loading"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3480
      Left            =   15
      TabIndex        =   3
      Top             =   1320
      Width           =   7365
      ExtentX         =   12991
      ExtentY         =   6138
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   750
      TabIndex        =   2
      Text            =   "http://www.bright.net/~stevensp"
      Top             =   900
      Width           =   6630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   930
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:05 PM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "8/9/97"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      Height          =   75
      Left            =   30
      Top             =   1185
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   915
      Width           =   690
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6645
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "brw1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "brw1.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "brw1.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "brw1.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "brw1.frx":0F72
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





     Private Sub Command1_Click()
         If Text1.Text <> "" Then
             WebBrowser1.Navigate Text1.Text
             If WebBrowser1.Visible = False Then
                 WebBrowser1.Visible = True
             End If
         End If
     End Sub

    



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1_Click
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "back"
WebBrowser1.GoBack
Case "forward"
WebBrowser1.GoForward
Case "Home"
Text1.Text = "http://www.bright.net/~stevensp"
Command1_Click
Case "refresh"
WebBrowser1.Refresh
Case "stop"
WebBrowser1.Stop
End Select



End Sub

Private Sub WebBrowser1_DownloadBegin()
StatusBar1.Panels(1).Text = " Getting Page..."
End Sub

Private Sub WebBrowser1_DownloadComplete()
StatusBar1.Panels(1).Text = "Got It"
End Sub
