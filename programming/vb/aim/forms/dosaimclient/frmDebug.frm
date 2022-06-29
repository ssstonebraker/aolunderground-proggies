VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmDebug 
   Caption         =   "AIM debug window"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfDebug 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmDebug.frx":1272
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  rtfDebug.SelColor = vbWhite
  rtfDebug.SelText = "Visual Basic 6 AOL Instant Messenger Example - (www.dosfx.com)"
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 2000 Then Me.Height = 2000
    rtfDebug.Width = Me.Width - 120
    rtfDebug.Height = Me.Height - 405
  End If
End Sub

