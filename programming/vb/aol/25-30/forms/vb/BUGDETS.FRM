VERSION 2.00
Begin Form frmBugs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Bug Detail Entry"
   ClientHeight    =   5340
   ClientLeft      =   2325
   ClientTop       =   2175
   ClientWidth     =   7365
   ControlBox      =   0   'False
   Height          =   5745
   Left            =   2265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7365
   Top             =   1830
   Width           =   7485
   Begin CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   5340
      TabIndex        =   8
      Top             =   4875
      Width           =   1890
   End
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   4875
      Width           =   1890
   End
   Begin PictureBox Picture2 
      Height          =   4260
      Left            =   120
      ScaleHeight     =   4230
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   525
      Width           =   7095
      Begin ComboBox cboFilename 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1395
         TabIndex        =   4
         Tag             =   "X 50"
         Top             =   1455
         Width           =   5205
      End
      Begin TextBox txtNotes 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2550
         Width           =   6870
      End
      Begin TextBox txtProcedure 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1395
         TabIndex        =   5
         Tag             =   "X 50"
         Top             =   2040
         Width           =   5295
      End
      Begin TextBox txtCreatedOn 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1395
         TabIndex        =   2
         Tag             =   "X 20"
         Top             =   765
         Width           =   2550
      End
      Begin TextBox txtDescription 
         BackColor       =   &H0000FFFF&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1395
         TabIndex        =   0
         Tag             =   "X 35"
         Top             =   150
         Width           =   5295
      End
      Begin CheckBox chkCleared 
         Caption         =   "Cleared"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4530
         TabIndex        =   3
         Top             =   765
         Width           =   1065
      End
      Begin Label Label7 
         Caption         =   "Procedure"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   105
         TabIndex        =   16
         Top             =   2115
         Width           =   1065
      End
      Begin Label Label6 
         Caption         =   "Filename"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   105
         TabIndex        =   15
         Top             =   1470
         Width           =   1005
      End
      Begin Label Label4 
         Caption         =   "Created on"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   105
         TabIndex        =   14
         Top             =   825
         Width           =   1245
      End
      Begin Label Label3 
         Caption         =   "Description"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   105
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin PictureBox Picture1 
      Height          =   450
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   7080
      TabIndex        =   9
      Top             =   60
      Width           =   7110
      Begin Label lblBugID 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5070
         TabIndex        =   12
         Top             =   60
         Width           =   1515
      End
      Begin Label lblSystem 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1035
         TabIndex        =   11
         Top             =   60
         Width           =   3840
      End
      Begin Label Label1 
         Caption         =   "&System"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "System"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   150
         TabIndex        =   10
         Top             =   75
         Width           =   795
      End
   End
End
Option Explicit

Sub cmdCancel_Click ()

    ' Set the global cancel flag to true
    gcancel = True

    frmBugs.Hide
    DoEvents


End Sub

Sub cmdOK_Click ()

    ' Set the global cancel variable to false
    gcancel = False
    frmBugs.Hide
    DoEvents

End Sub

Sub Form_Load ()

    ' Centre the form on the screen
    frmBugs.Left = (Screen.Width - frmBugs.Width) / 2
    frmBugs.Top = (Screen.Height - frmBugs.Height) / 2

End Sub

Sub txtCreatedOn_GotFocus ()

    Call GotFocus(txtCreatedOn)

End Sub

Sub txtCreatedOn_KeyPress (KeyAscii As Integer)

    Call KeyCheck(txtCreatedOn, KeyAscii)

End Sub

Sub txtDescription_GotFocus ()

    Call GotFocus(txtDescription)

End Sub

Sub txtDescription_KeyPress (KeyAscii As Integer)

    Call KeyCheck(txtDescription, KeyAscii)

End Sub

Sub txtNotes_GotFocus ()

    Call GotFocus(txtNotes)

End Sub

Sub txtProcedure_GotFocus ()

    Call GotFocus(txtProcedure)

End Sub

Sub txtProcedure_KeyPress (KeyAscii As Integer)

    Call KeyCheck(txtProcedure, KeyAscii)

End Sub

