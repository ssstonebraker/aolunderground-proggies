VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShellExecute Demo Project"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWebSite 
      Caption         =   "Open VB-World in the default web browser"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open this File with the default program"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblJelsoft 
      Caption         =   $"frmMain.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblFile 
      Caption         =   "&File to Open"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdOpen_Click()

'check that the file in the text box exists
If Dir(txtFile) = "" Then
    Call MsgBox("The file in the text box does not exist.", vbExclamation)
    Exit Sub
End If

'open the file with the default program
Call ShellExecute(hwnd, "Open", txtFile, "", App.Path, 1)

'Note:  This is the equivalent of
'right clicking on a file in Windows95
'and selecting "Open"
'
'If you would like to do something
'else to the file rather than opening
'it, right click on a file and see
'what options are in the menu.  Then
'change the "Open" in the code above to
'read what the menu item says.
'
'Note:  This code is great for opening
'web document into the default
'browser.  To open http://www.jelsoft.com
'into the default browser the following
'code would be used:
'
'Call ShellExecute(hwnd,"Open","http://www.jelsoft.com","",app.path,1)
'
'For more demos, please visit Jelsoft VB-World at
'http://www.jelsoft.com
'
'If you have a question or a query, please
'send an email to vbw@jelsoft.com.

End Sub

Private Sub cmdWebSite_Click()
'open up VB-World in the default browser.
Call ShellExecute(hwnd, "Open", "http://www.jelsoft.com/vbw/", "", App.Path, 1)
End Sub
