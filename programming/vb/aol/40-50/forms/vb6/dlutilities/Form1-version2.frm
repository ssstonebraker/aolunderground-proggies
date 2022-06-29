VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   5670
   ClientWidth     =   4320
   Icon            =   "Form1-version2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1-version2.frx":030A
   ScaleHeight     =   3225
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   360
      Top             =   3120
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   4080
      X2              =   4080
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   4080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "options/misc"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MousePointer    =   10  'Up Arrow
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "stop auto sign on"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "start auto sign on"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   4080
      X2              =   4080
      Y1              =   600
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   4080
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   240
      Y1              =   600
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   4080
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormAbove Me
mIRCSendChat ("•.·—› download utilities v2 by koko")
AOL40SendChat ("•.·—› download utilities v2 by koko")
TimeOut (0.3)
mIRCSendChat ("•.·—› loaded by: ") + GetUser
AOL40SendChat ("•.·—› loaded by: ") + GetUser
TimeOut (0.3)
mIRCSendChat ("•.·—› get it at http://www.angelfire.com/yt/koko")
AOL40SendChat ("< A HREF=") + "http://www.angelfire.com/yt/koko" + (">") + ("•.·—›") + (" http://www.angelfire.com/yt/koko") + ("</A>")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Label1_Click()

Dim strAdd As String
strAdd = InputBox("Please enter your password, this if for when you get logged off it will log back on and continue the download.", "Please enter your AOL password.")

If TrimSpaces(strAdd) = "" Then
Exit Sub
Else
mIRCSendChat ("•.·—› download utilities v2 by koko")
AOL40SendChat ("•.·—› download utilities v2 by koko")
TimeOut (0.3)
mIRCSendChat ("•.·—› auto sign on is on!")
AOL40SendChat ("•.·—› auto sign on is on!")
TimeOut (0.3)
mIRCSendChat ("•.·—› get it at http://www.angelfire.com/yt/koko")
AOL40SendChat ("< A HREF=") + "http://www.angelfire.com/yt/koko" + (">") + ("•.·—›") + (" http://www.angelfire.com/yt/koko") + ("</A>")

Label4.Caption = (strAdd)
Timer1.Enabled = True
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80FF&
End Sub

Private Sub Label2_Click()
mIRCSendChat ("•.·—› download utilities v2 by koko")
AOL40SendChat ("•.·—› download utilities v2 by koko")
TimeOut (0.3)
mIRCSendChat ("•.·—› auto sign on is off!")
AOL40SendChat ("•.·—› auto sign on is off!")
TimeOut (0.3)
mIRCSendChat ("•.·—› get it at http://www.angelfire.com/yt/koko")
AOL40SendChat ("< A HREF=") + "http://www.angelfire.com/yt/koko" + (">") + ("•.·—›") + (" http://www.angelfire.com/yt/koko") + ("</A>")
Timer1.Enabled = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80FF&
End Sub

Private Sub Label3_Click()
Form2.PopupMenu Form2.options
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H80FF&
End Sub

Private Sub Timer1_Timer()
If GetUser = "" Then
Call AOL40SignOnWithPW(Label4.Caption)
TimeOut (60)
If Form2.Label1.Caption = LCase("y") Then
AOL40ContinueDownload
TimeOut (15)
AOL40Keyword ("aol://2719:2-2-") + Form2.Label2.Caption
End If
End If
End Sub
