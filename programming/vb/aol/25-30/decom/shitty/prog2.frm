VERSION 4.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AOL Directory"
   ClientHeight    =   4230
   ClientLeft      =   2715
   ClientTop       =   1785
   ClientWidth     =   4185
   ForeColor       =   &H00000000&
   Height          =   4635
   Left            =   2655
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Top             =   1440
   Width           =   4305
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Detect"
      Height          =   225
      Left            =   1830
      TabIndex        =   5
      Top             =   3630
      Width           =   1395
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   3540
      Width           =   1605
   End
   Begin VB.Timer Timer1 
      Left            =   3390
      Top             =   4320
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   3180
      Left            =   1800
      TabIndex        =   3
      Top             =   300
      Width           =   2325
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   3180
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   225
      Left            =   3420
      TabIndex        =   1
      Top             =   3630
      Width           =   675
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   3930
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter in the path which aol is in"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   990
      TabIndex        =   0
      Top             =   30
      Width           =   2265
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form4.Show
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Click this once you have auto or manually found AOL"
End Sub


Private Sub Command2_Click()
mypath = "c:\"
Myname = Dir(mypath, vbDirectory)
Do While Myname <> ""
If Myname <> "." And Myname <> ".." Then
    If (GetAttr(mypath & Myname) And vbDirectory) = vbDirectory Then
    mark$ = Myname
    Bob = Mid(mark$, 1, 3)
        If Bob = "AOL" Then
            Msg = "Aol has been detected in the Directory (C:\" & Myname & "). Would you like to use this directory for this programs AOL directory default.?"
            Style = vbYesNo
            Title = "AOL Detection"
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then
                mypath = CurDir
                Open mypath & "\deicide.ini" For Random As #1
                Path$ = "C:\" & Myname & "\Waol.exe"
                Put #1, 1, Path$
                MsgBox "Ok, everythings done...your ready to go. Have fun! :)"
                Close #1
                Exit Sub
            End If
        End If
    End If
End If
Myname = Dir
Loop
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "If you choose this way it auto detects AOL...hopefully :)"
End Sub


Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Select a folder to look in. ie(AOL30A)"
End Sub


Private Sub File1_DblClick()
If File1 = "waol.exe" Or File1 = "aol.exe" Then
mypath = CurDir
Open mypath & "\deicide.ini" For Random As #1
Path$ = File1.Path & "\" & File1
Put #1, 1, Path$
MsgBox "Ok, your ready to go! :)"
Close #1
Else
MsgBox "Look for either waol.exe or aol.exe it won't work otherwise sorry"
End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Double click a file to select that as AOL. ie(Waol.exe)"
End Sub


Private Sub Form_Load()
Call StayOnTop(Form2)
Timer1.interval = 5

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "This does nothing..."
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "Another fun little box"
End Sub


Private Sub Timer1_Timer()
Dir1.Path = Drive1
File1.Path = Dir1
End Sub


