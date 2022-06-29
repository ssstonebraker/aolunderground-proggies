VERSION 2.00
Begin Form Fisher 
   BackColor       =   &H8000000F&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Password Fisher"
   ClientHeight    =   4860
   ClientLeft      =   4875
   ClientTop       =   345
   ClientWidth     =   4635
   Height          =   5265
   HelpContextID   =   90
   Left            =   4815
   LinkTopic       =   "Form2"
   ScaleHeight     =   4860
   ScaleWidth      =   4635
   Top             =   0
   Width           =   4755
   Begin SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BorderWidth     =   4
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   8535
      Begin SSCommand Command3D6 
         Caption         =   "E&xit"
         Font3D          =   1  'Raised w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   3240
         TabIndex        =   18
         Top             =   3720
         Width           =   975
      End
      Begin SSCommand Command3D5 
         Caption         =   "&Stop"
         Font3D          =   1  'Raised w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   3720
         Width           =   975
      End
      Begin SSCommand Command3D4 
         Caption         =   "&Cast away"
         Font3D          =   1  'Raised w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2400
         Picture         =   FISHER.FRX:0000
         ScaleHeight     =   615
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   2880
         Width           =   495
      End
      Begin Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1400
         Left            =   360
         Top             =   3720
      End
      Begin Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   480
         Top             =   4440
      End
      Begin SSFrame Frame3D3 
         Alignment       =   2  'Center
         Caption         =   "# Fish's Sent"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   3000
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
         Begin PictureBox Input2 
            BackColor       =   &H00FFFFFF&
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   10.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   360
            MousePointer    =   1  'Arrow
            ScaleHeight     =   345
            ScaleWidth      =   825
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
      End
      Begin SSFrame Frame3D2 
         Alignment       =   2  'Center
         Caption         =   "# of people in room"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
         Begin TextBox Text2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
      End
      Begin SSFrame Frame3D1 
         Alignment       =   2  'Center
         Caption         =   "Fishing Sayings"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   9
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   2415
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         Begin SSOption Option3D7 
            Caption         =   "6"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   9
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin SSOption Option3D6 
            Caption         =   "5"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   9
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2400
            Width           =   495
         End
         Begin SSOption Option3D5 
            Caption         =   "4"
            Font3D          =   3  'Inset w/light shading
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   9
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin TextBox Text1 
            Enabled         =   0   'False
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000080&
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   600
            Width           =   4095
         End
         Begin SSOption Option3D4 
            Caption         =   "&Use your own"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   12
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2040
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin SSOption Option3D3 
            Caption         =   "3"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
         End
         Begin SSOption Option3D2 
            Caption         =   "2"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   840
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
         End
         Begin SSOption Option3D1 
            Caption         =   "1"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "Times New Roman"
            FontSize        =   12
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
         End
      End
      Begin Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Menace II AOL"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   24
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800080&
         Height          =   735
         Left            =   960
         TabIndex        =   19
         Top             =   4320
         Width           =   3615
      End
      Begin Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000F&
         BackStyle       =   0  'Transparent
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   36
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   870
         Left            =   4800
         TabIndex        =   13
         Top             =   960
         Width           =   180
      End
   End
End
Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer



Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Sub Command3D1_Click ()
    If text2 = "" Then
        MsgBox "Gotta send at least 1 fish", 16, "Dumbass!"
    Else
        Clipboard.Clear
        Clipboard.SetText text1.Text
        AppActivate "America  Online"
        timer1.Enabled = True
    End If

End Sub

Sub Command3D2_Click ()
timer1.Enabled = False
timer2.Enabled = False
input2 = "0"
input1 = "0"
End Sub

Sub Command3D3_Click ()
Unload Me


End Sub

Sub Command3D4_Click ()
    If text2 = "" Then
        MsgBox "Gotta send at least 1 fish", 16, "Dumbass!"
    Else
        Clipboard.Clear
        Clipboard.SetText text1.Text
        AppActivate "America  Online"
        timer1.Enabled = True
    End If

End Sub

Sub Command3D5_Click ()
timer1.Enabled = False
timer2.Enabled = False
input2 = "0"
input1 = "0"

End Sub

Sub Form_Resize ()
    Dim Success%
   
  Success% = SetWindowPos%(Fisher.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) '(all on
End Sub

Sub Option3D1_Click (Value As Integer)
text1.Text = "Hello, and welcome, to the world of America Online™, due to a sudden system failure, our new member database has been formatted, please click on 'respond' and send me your personal log on password, thank you, and enjoy the service !"

End Sub

Sub Option3D2_Click (Value As Integer)
    text1.Text = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
End Sub

Sub Option3D3_Click (Value As Integer)
    text1.Text = "Hello and welcome to America Online!  Because portions of your account information has been inadvertantly destroyed, we must ask you to click on 'respond' and send us your personal logon password immediately.  Thank you for your cooperation and enjoy the service!"
End Sub

Sub Option3D4_Click (Value As Integer)
text1.Text = ""

End Sub

Sub Option3D5_Click (Value As Integer)
text1.Text = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."

End Sub

Sub Option3D6_Click (Value As Integer)
text1.Text = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."


End Sub

Sub Option3D7_Click (Value As Integer)
text1.Text = "I work for the America On-Line Billing Department. Welcome to America Online!  Because portions of your account information has been inadvertantly destroyed, we must ask you to click on RESPOND and send us your personal logon password immediately.  Thank you for your cooperation and enjoy the service... America On-Line Billing Department."


End Sub

Sub Timer1_Timer ()
    If input1 = input2 Then
        timer1.Enabled = False
        timer2.Enabled = False
        MsgBox "The Fisher is finished.", 16, "End"
        input2 = "0"
    Else
        SendKeys "{ENTER}"
        DoEvents
        SendKeys "^V"
        SendKeys "{TAB}"
        SendKeys "{ }"
        DoEvents
        timer1.Enabled = False
        timer2.Enabled = True
        input2 = input2 + 1
    End If

End Sub

Sub Timer2_Timer ()
    SendKeys "{DOWN}"
    timer1.Enabled = True
    timer2.Enabled = False

End Sub

