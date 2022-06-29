VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C548D53D-4DB1-11D2-A11D-549F06C10000}#1.0#0"; "FileEncryption.ocx"
Begin VB.Form frmEncryptTest 
   Caption         =   "File Encryption Test"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin FileEncryption.FileEncryptor FileEncryptor1 
      Left            =   1080
      Top             =   0
      _ExtentX        =   3731
      _ExtentY        =   1217
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1200
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "txtFile"
      Top             =   390
      Width           =   3855
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1200
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1200
   End
   Begin VB.TextBox txtPassword1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtPassword2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog cdlOne 
      Left            =   2925
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17491e-38
   End
   Begin VB.Label lblFile 
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1755
   End
   Begin VB.Label lblPassword1 
      Caption         =   "Enter password:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblPassword2 
      Caption         =   "Enter password again:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "frmEncryptTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    cdlOne.DialogTitle = "File Encryptor 1.0"
    cdlOne.Flags = cdlOFNHideReadOnly
    cdlOne.Filter = "All files (*.*)|*.*"
    cdlOne.CancelError = True
    On Error Resume Next
    cdlOne.ShowOpen
    If Err = 0 Then
        txtFile.Text = cdlOne.FileName
    End If
    On Error GoTo 0
End Sub

Private Sub cmdEncrypt_Click()
    If txtPassword1.Text <> txtPassword2.Text Then
        MsgBox "The two passwords are not the same!", vbExclamation, "File Encryptor 1.0"
        Exit Sub
    End If
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdBrowse.Enabled = False
    Refresh
    FileEncryptor1.Encrypt txtFile.Text, txtPassword1.Text
    txtFile_Change
    MousePointer = vbDefault
End Sub

Private Sub cmdDecrypt_Click()
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdBrowse.Enabled = False
    Refresh
    FileEncryptor1.Decrypt txtFile.Text, txtPassword1.Text
    txtFile_Change
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    txtFile.Text = ""
End Sub

Private Sub txtFile_Change()
    Dim lngFileLen As Long
    Dim strHead As String
    'Check to see whether file exists
    On Error Resume Next
    lngFileLen = Len(Dir(txtFile.Text))
    'Disable buttons if filename isn't valid
    If Err <> 0 Or lngFileLen = 0 Or Len(txtFile.Text) = 0 Then
        cmdEncrypt.Enabled = False
        cmdDecrypt.Enabled = False
        lblPassword1.Enabled = False
        txtPassword1.Enabled = False
        lblPassword2.Enabled = False
        txtPassword2.Enabled = False
        txtPassword2.Text = ""
        Exit Sub
    End If
    'Get first 8 bytes of selected file
    Open txtFile.Text For Binary As #1
    strHead = Space(8)
    Get #1, , strHead
    Close #1
    'Check to see whether file is already encrypted
    If strHead = "[Secret]" Then
        cmdEncrypt.Enabled = False
        cmdDecrypt.Enabled = True
        lblPassword2.Enabled = False
        txtPassword2.Enabled = False
        txtPassword2.Text = ""
    Else
        cmdEncrypt.Enabled = True
        cmdDecrypt.Enabled = False
        lblPassword2.Enabled = True
        txtPassword2.Enabled = True
    End If
    lblPassword1.Enabled = True
    txtPassword1.Enabled = True
    cmdBrowse.Enabled = True
End Sub
