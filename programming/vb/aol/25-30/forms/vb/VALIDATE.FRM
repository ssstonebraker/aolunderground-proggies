VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Chapter 13 - Object Variables"
   ClientHeight    =   2340
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5325
   Height          =   2745
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5325
   Top             =   1140
   Width           =   5445
   Begin PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   1515
      Left            =   255
      ScaleHeight     =   1485
      ScaleWidth      =   4830
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   585
      Width           =   4860
      Begin TextBox txtText 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Tag             =   "A12"
         Top             =   120
         Width           =   2895
      End
      Begin TextBox txtText 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Tag             =   "94"
         Top             =   525
         Width           =   2895
      End
      Begin TextBox txtText 
         Height          =   285
         Index           =   2
         Left            =   1785
         TabIndex        =   2
         Tag             =   "X5"
         Top             =   945
         Width           =   2925
      End
      Begin Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Anything at all"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   990
         Width           =   1320
      End
      Begin Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Numbers Only"
         Height          =   240
         Left            =   195
         TabIndex        =   6
         Top             =   555
         Width           =   1425
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Alphabetic Only"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   150
         Width           =   1410
      End
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Validation with Object Variables"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   255
      TabIndex        =   4
      Top             =   90
      Width           =   4860
   End
End
Option Explicit

Sub txtText_GotFocus (Index As Integer)

    txtText(Index).SelStart = 0
    txtText(Index).SelLength = Len(txtText(Index).Text)

End Sub

Sub txtText_KeyPress (Index As Integer, KeyAscii As Integer)

    Call ValidateKeyPress(txtText(Index), KeyAscii)

End Sub

Sub ValidateKeyPress (txtControl As TextBox, nKeyAscii As Integer)

    Dim sMaxLength As String
    Dim sKey As String * 1

    ' first we need to ignore the special keys, such as Delete, or the cursor keys
    If nKeyAscii < 32 Or nKeyAscii > 126 Then Exit Sub

    ' Now we need to get the maximum length allowed in the text box
    sMaxLength = Right(txtControl.Tag, Len(txtControl.Tag) - 1)

    ' Since we know the maximum length of data allowed we can check what is in the text box
    ' and throw the keypress out if it means going over the limit.
    If Len(txtControl.Text) = Val(sMaxLength) Then
        Beep
        nKeyAscii = 0
        Exit Sub
    End If


    ' Assuming that the length of data will not be exceed, does the keypress fit in with the
    ' kind of data allowed, ie numeric, alphabetic or anything else
    Select Case Left$(txtControl.Tag, 1)
        
        ' If the textbox only allows A-Z and a-z then...
        Case "A"
            ' First convert the keypress to Uppercase; this simplifies the testing process
            sKey = UCase(Chr$(nKeyAscii))
            
            ' If the keypressed is outside A to Z then throw it away
            If Asc(sKey) < 65 Or Asc(sKey) > 90 Then
                Beep
                nKeyAscii = 0
                Exit Sub
            End If

        ' If the textbox only allows 0-9 then...
        Case "9"
            If nKeyAscii < 48 Or nKeyAscii > 57 Then
                Beep
                nKeyAscii = 0
                Exit Sub
            End If

    End Select

End Sub

