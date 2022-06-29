VERSION 5.00
Begin VB.Form SRC14 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SRC14"
   ClientHeight    =   1260
   ClientLeft      =   1140
   ClientTop       =   1485
   ClientWidth     =   5175
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1260
   ScaleWidth      =   5175
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Reboot System!"
      Height          =   525
      Left            =   390
      TabIndex        =   0
      Top             =   270
      Width           =   4455
   End
End
Attribute VB_Name = "SRC14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer

Private Sub Command1_Click()
    Dim x As Integer
    If MsgBox("This routine will Reboot the Computer!" & Chr$(13) & Chr$(10) & "Are you sure you want to do this now?", 36, "Restart Windows?") = 6 Then
        x = RebootSystem()
        If Not x Then MsgBox "Some program(s) refused to terminate", 48, "KnK Demo"
    End If
End Sub

Private Function RebootSystem() As Integer
'   Causes the computer to be rebooted
'   If any programs refuse to terminate, then this function will return a ZERO
    Dim i As Integer
    Dim EW_REBOOTSYSTEM As Long
    EW_REBOOTSYSTEM = &H43
    i = ExitWindows(EW_REBOOTSYSTEM, 0)
End Function

