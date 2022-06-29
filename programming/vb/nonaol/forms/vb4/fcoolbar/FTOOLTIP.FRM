VERSION 4.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   390
   ClientLeft      =   3060
   ClientTop       =   4455
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   795
   Left            =   3000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Top             =   4110
   Visible         =   0   'False
   Width           =   3345
   Begin VB.Label lblToolTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblTip"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   375
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' This form shows a tool tip
' to show it set frmToolTip.lblToolTip.Caption to a non null value
' to hide it again set it to an empty string

Const SW_SHOWNA = 8
Const SW_HIDE = 0

#If Win16 Then
    Private Declare Function ShowWindow Lib "user" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
#End If

Private Sub lblToolTip_Change()

    ' the caption in the label control has changed
    ' resize the form to make sure it is big enough
    #If Win16 Then
        Dim iRet As Integer
    #Else
        Dim iRet As Long
    #End If

    If Len(Trim$(lblToolTip.Caption)) > 0 Then

        ' make sure the form is big enough
        lblToolTip.Top = 15
        lblToolTip.Left = 15
        Width = lblToolTip.Left + lblToolTip.Width + 45
        Height = lblToolTip.Top + lblToolTip.Height + 45
    
        ' make sure all of the form is visible
        If Left + Width > Screen.Width Then
            Left = Screen.Width - Width
        End If
        If Top + Height > Screen.Height Then
            Top = Screen.Height - Height
        End If
        If Left < 0 Then
            Left = 0
        End If
        If Top < 0 Then
            Top = 0
        End If

        ' show the form without giving it the focus
        iRet = ShowWindow(hWnd, SW_SHOWNA)
        
        ' NOTE: The Forms BorderStyle property is set to 0 to prevent it from
        ' appearing on the taskbar when activated - so we'll have to draw a
        ' border around it manually
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), , B
    Else
        
        ' the label caption has been changed to an empty string
        ' unload the form

        iRet = ShowWindow(hWnd, SW_HIDE)
        Unload Me
        Set frmToolTip = Nothing
    End If

End Sub

