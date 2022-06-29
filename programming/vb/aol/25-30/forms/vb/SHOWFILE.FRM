VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Testing multiple conditions"
   ClientHeight    =   1725
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7005
   Height          =   2130
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   7005
   Top             =   1140
   Width           =   7125
   Begin PictureBox Picture1 
      Height          =   1575
      Left            =   90
      ScaleHeight     =   1545
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   75
      Width           =   6795
      Begin CommandButton cmdView 
         Caption         =   "View The File"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   15
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   660
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   6285
      End
      Begin CommandButton cmdSelectFile 
         Caption         =   "File Name"
         Height          =   375
         Left            =   5325
         TabIndex        =   3
         Top             =   105
         Width           =   1215
      End
      Begin TextBox txtFile 
         BackColor       =   &H00FFFF00&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   105
         Width           =   3570
      End
      Begin Label Label1 
         Caption         =   "File to view"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   135
         Width           =   1455
      End
   End
   Begin CommonDialog dlgFileName 
      Left            =   705
      Top             =   4635
   End
End
Option Explicit

Sub cmdSelectFile_Click ()

    ' Turn on the error handling - when the user hits the Cancel button in the File dialog this
    ' will trigger a run time error, which we don't want to halt the program.
    On Error GoTo Dialog_Error

    ' Set the common dialog control's CancelError property to True, meaning that if the user
    ' clicks on the Cancel button then a run time error is generated
    dlgFileName.CancelError = True

    ' Set up the common dialog's title
    dlgFileName.DialogTitle = "Select file to view"

    ' Set up the filters to show a list of all files (*.*)
    dlgFileName.Filter = "All Files | *.*"
    dlgFileName.FilterIndex = 1

    ' Setting the Action property to 1 causes the dialog box to do its "thang"
    dlgFileName.Action = 1

    ' Copy the selected filename into the textbox control.
    txtFile.Text = dlgFileName.Filename

Dialog_Error:
    On Error GoTo 0
    Exit Sub

End Sub

Sub cmdView_Click ()
    
    ' This sets up a variable, a place in memory to hold some data.
    ' The As String bit tells VB that we are going to store Text in this area.
    Dim sExtension As String

    Dim ReturnValue

    sExtension = UCase(Right$(txtFile, 3))

    If Dir$(txtFile.Text) = "" Then
        MsgBox "Sorry, I couldnt find that file!"
        Exit Sub
    
    ElseIf sExtension = "TXT" Then
        
        ReturnValue = Shell("Notepad " & txtFile.Text, 1)

    ElseIf sExtension = "WRI" Then

        ReturnValue = Shell("Write " & txtFile.Text, 1)

    ElseIf sExtension = "CRD" Then

        ReturnValue = Shell("Cardfile " & txtFile.Text, 1)

    End If
    
End Sub

