VERSION 2.00
Begin Form frmSystem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "System Details Entry"
   ClientHeight    =   4485
   ClientLeft      =   1470
   ClientTop       =   2055
   ClientWidth     =   7365
   ControlBox      =   0   'False
   Height          =   4890
   Left            =   1410
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7365
   Top             =   1710
   Width           =   7485
   Begin CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   5640
      TabIndex        =   5
      Top             =   3900
      Width           =   1650
   End
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   3915
      Width           =   1650
   End
   Begin CommonDialog dlgFiles 
      DialogTitle     =   "Select a project for this system"
      Filter          =   "Mak files | *.mak"
      FilterIndex     =   1
      Left            =   3435
      Top             =   2205
   End
   Begin PictureBox Picture1 
      Height          =   3765
      Left            =   105
      ScaleHeight     =   3735
      ScaleWidth      =   7125
      TabIndex        =   6
      Top             =   90
      Width           =   7155
      Begin TextBox txtNotes 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2145
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1455
         Width           =   6840
      End
      Begin CommandButton cmdFile 
         Caption         =   "&File"
         Height          =   375
         Left            =   6120
         TabIndex        =   2
         Top             =   690
         Width           =   825
      End
      Begin TextBox txtMakefile 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1755
         TabIndex        =   1
         Tag             =   "X 50"
         Top             =   750
         Width           =   4305
      End
      Begin TextBox txtSystem 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1740
         TabIndex        =   0
         Tag             =   "X 35"
         Top             =   90
         Width           =   5145
      End
      Begin Label Label3 
         Caption         =   "Project Name"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   165
         TabIndex        =   8
         Top             =   765
         Width           =   1425
      End
      Begin Label Label1 
         Caption         =   "System Name"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   165
         TabIndex        =   7
         Top             =   135
         Width           =   1485
      End
   End
End
Option Explicit

Sub cmdCancel_Click ()

    gCancel = True
    frmSystem.Hide
    DoEvents

End Sub

Sub cmdFile_Click ()

    dlgFiles.Action = 1

    txtMakeFile.Text = dlgFiles.Filename

End Sub

Sub cmdOK_Click ()

    gCancel = False
    frmSystem.Hide
    DoEvents


End Sub

Sub txtMakefile_GotFocus ()
    
    Call GotFocus(txtMakeFile)

End Sub

Sub txtMakefile_KeyPress (Keyascii As Integer)

    Call KeyCheck(txtMakeFile, Keyascii)

End Sub

Sub txtNotes_GotFocus ()

    Call GotFocus(txtNotes)

End Sub

Sub txtNotes_KeyPress (Keyascii As Integer)

    Call KeyCheck(txtNotes, Keyascii)

End Sub

Sub txtSystem_GotFocus ()

    Call GotFocus(txtSystem)

End Sub

Sub txtSystem_KeyPress (Keyascii As Integer)

    Call KeyCheck(txtSystem, Keyascii)
    
End Sub

