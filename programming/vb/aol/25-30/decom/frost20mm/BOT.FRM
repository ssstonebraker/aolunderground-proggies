VERSION 2.00
Begin Form bot 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Mass Mail Bot"
   ClientHeight    =   3300
   ClientLeft      =   3096
   ClientTop       =   1596
   ClientWidth     =   3144
   Height          =   3636
   Left            =   3048
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3144
   Top             =   1308
   Width           =   3240
   Begin SSFrame Frame3D3 
      Alignment       =   2  'Center
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin SSCommand Command3D3 
         Caption         =   "Exit"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   2760
         Width           =   975
      End
      Begin SSCommand Command3D2 
         Caption         =   "Stop Bot"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin SSCommand Command3D1 
         Caption         =   "Start Bot"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin CommandButton Command3 
         Caption         =   "?"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   2760
         TabIndex        =   19
         Top             =   1650
         Width           =   255
      End
      Begin TextBox Text3 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Text            =   "5"
         Top             =   1920
         Width           =   495
      End
      Begin CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enable Manager"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1455
      End
      Begin CommandButton Command8 
         Caption         =   "Clear"
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   2160
         TabIndex        =   15
         Top             =   1650
         Width           =   615
      End
      Begin CommandButton Command6 
         Caption         =   "Add"
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   1680
         TabIndex        =   14
         Top             =   1650
         Width           =   495
      End
      Begin TextBox Text1 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin ListBox List2 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   870
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin CommandButton Command5 
         Caption         =   "Clear"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   1650
         Width           =   495
      End
      Begin CommandButton Command4 
         Caption         =   "Room"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   1650
         Width           =   615
      End
      Begin CommandButton Command2 
         Caption         =   "+"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1650
         Width           =   255
      End
      Begin ListBox List1 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VScrollBar VScroll1 
         Height          =   350
         Left            =   990
         Max             =   32000
         TabIndex        =   5
         Top             =   2350
         Value           =   15
         Width           =   255
      End
      Begin TextBox Text2 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "FrostBite v2.o"
         Top             =   2350
         Width           =   1215
      End
      Begin Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   0
         Top             =   3960
      End
      Begin VBMsg VBMsg1 
         Height          =   420
         Left            =   0
         MessageCount    =   BOT.FRX:0000
         MessageList     =   BOT.FRX:0002
         MessageTypes    =   0  'Selected Messages
         PostDefault     =   0   'False
         Top             =   3960
         Width           =   420
      End
      Begin Line Line1 
         X1              =   120
         X2              =   3000
         Y1              =   2280
         Y2              =   2280
      End
      Begin Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Minute(s):"
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1940
         Width           =   735
      End
      Begin Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Room Manager"
         FontBold        =   0   'False
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "People: 0"
         FontBold        =   0   'False
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2400
         Width           =   495
      End
      Begin Label Label2 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   615
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "On:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   2400
         Width           =   375
      End
   End
End
' BOT.FRM
Option Explicit
Dim m001A As Variant
Dim m002A As Integer
Dim m002C As Variant
Const mc00D2 = 12 ' &HC%
Const mc00DA = 258 ' &H102%
Const mc00E0 = 16 ' &H10%

Sub Check1_Click ()
If  Check1.Value <> False Then
List2.Enabled = True
Text1.Enabled = True
Text3.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Label6.Enabled = True
If  Text3.Text <> "" Then
Label3.Caption = (Text3.Text * List2.ListCount) + Text3.Text
VScroll1.Value = (Text3.Text * List2.ListCount) + Text3.Text
End If
ElseIf Check1.Value = False Then
List2.Enabled = False
Text1.Enabled = False
Text3.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Label6.Enabled = False
End If
End Sub

Sub Command2_Click ()
Dim l0072 As Variant
Dim l0076 As Variant
l0072 = InputBox("Please enter the username to add", "Add")
If  l0072 = "" Then Exit Sub
For l0076 = 0 To List1.ListCount - 1
    If  UCase$(List1.List(l0076)) = UCase$(l0072) Then Exit Sub
Next l0076
List1.AddItem l0072
Label4.Caption = "People: " + List1.ListCount
End Sub

Sub Command3_Click ()
MsgBox "To remove something from a listbox just double click on the room or name you want removed."
End Sub

Sub Command3D1_Click ()
Dim l0096 As Integer
If  Text2.Text = "" Then MsgBox "You need to enter a phrase to get on", 16, "Error": Exit Sub
If  Label3.Caption = "0" Then MsgBox "Please enter some time to run bot!", 16, "Error": Exit Sub
Command3D2.Enabled = True
m002A% = Label3.Caption - Text3.Text
m001A = 0
l0096% = extfn0B24(fn03A3(), "_AOL_View")
VBMsg1.SubClasshWnd = l0096%
sub055F "×–•[·FrostBite·]•–× FrostBite Mass Mailer v2.o"
sub060A .001#
sub055F "×–•[·FrostBite·]•–× Mass Mail Going Out in: " + Label3.Caption + ""
sub060A .001#
sub055F "×–•[·FrostBite·]•–× Say """ + Text2 + """ To Get On"
Label3.Caption = Val(Label3.Caption) - 1
Timer1.Enabled = True
Command3D1.Enabled = False
End Sub

Sub Command3D2_Click ()
Command3D1.Enabled = True
sub055F "×–•[·FrostBite·]•–× The Mass Mail Has Stopped!"
m001A = 1
VBMsg1.SubClasshWnd = 0
Timer1.Enabled = False
Command3D2.Enabled = False
End Sub

Sub Command3D3_Click ()
MMMain.Show
Unload Me
End Sub

Sub Command4_Click ()
Dim l00AE As Integer
l00AE% = fn03A3()
Call sub01DD(l00AE%, List1)
Label4.Caption = "People: " + List1.ListCount
End Sub

Sub Command5_Click ()
List1.Clear
Label4.Caption = "People: 0"
End Sub

Sub Command6_Click ()
Dim l00B0 As Variant
If  Text1.Text = "" Then MsgBox "Please enter a room", 16, "Error": Exit Sub
List2.AddItem Text1.Text
l00B0 = ((List2.ListCount) * (Text3.Text))
Label3.Caption = l00B0 + Text3.Text
VScroll1.Value = l00B0 + Text3.Text
Text1.Text = ""
End Sub

Sub Command8_Click ()
List2.Clear
End Sub

Sub Form_Load ()
sub05BC Me
Me.Top = 0
Me.Left = 0
List2.AddItem "fate"
List2.AddItem "fate2"
List2.AddItem "fate3"
List2.AddItem "fate4"
List2.AddItem "fate5"
List2.AddItem "fate6"
List2.AddItem "fate7"
List2.AddItem "fate8"
List2.AddItem "fate9"
List2.AddItem "poa"
List2.AddItem "poa2"
List2.AddItem "poa3"
List2.AddItem "poa4"
List2.AddItem "poa5"
List2.AddItem "havok"
List2.AddItem "havok2"
List2.AddItem "havok3"
List2.AddItem "server"
List2.AddItem "server2"
List2.AddItem "server3"
List2.AddItem "server4"
List2.AddItem "server5"
List2.AddItem "server6"
List2.AddItem "server7"
End Sub

Sub Form_Unload (Cancel As Integer)
Dim l00B6 As Variant
MMMain.List1.Clear
For l00B6 = 0 To List1.ListCount - 1
MMMain.List1.AddItem List1.List(l00B6)
Next l00B6
MMMain.Frame3D1.Caption = "People " + MMMain.List1.ListCount
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
Label4.Caption = "People: " & List1.ListCount
End Sub

Sub List2_DblClick ()
List2.RemoveItem List2.ListIndex
Label3.Caption = Label3.Caption - Text3.Text
VScroll1.Value = VScroll1.Value - Text3.Text
End Sub

Sub Text1_KeyPress (KeyAscii As Integer)
Dim l00BC As Variant
If  KeyAscii = 13 Then
If  Text1.Text = "" Then MsgBox "Please enter a room", 16, "Error": Exit Sub
List2.AddItem Text1.Text
l00BC = ((List2.ListCount) * (Text3.Text))
Label3.Caption = l00BC + Text3.Text
VScroll1.Value = l00BC + Text3.Text
Text1.Text = ""
End If
End Sub

Sub Text3_Change ()
If  Text3.Text <> "" Then
Label3.Caption = (Text3.Text * List2.ListCount) + Text3.Text
VScroll1.Value = (Text3.Text * List2.ListCount) + Text3.Text
End If
End Sub

Sub Timer1_Timer ()
Dim l00C0 As Variant
Dim l00C6 As String
Dim l00CA As Integer
Dim l00CE As Integer
Dim l00D0 As Integer
' Const mc00D2 = 12 ' &HC%
Dim l00D6 As Variant
' Const mc00DA = 258 ' &H102%
Dim l00DE As Integer
' Const mc00E0 = 16 ' &H10%
If  Check1.Value = 1 Then
    If  List2.ListCount <> 0 And Label3.Caption = m002A% Then
            For l00C0 = 0 To List2.ListCount - List2.ListCount
            List2.Selected(l00C0) = True
            Next l00C0
    sub055F "×–•[·FrostBite·]•–× Switching To: " & List2.Text & ""
    m002C = 1
    l00C6$ = fn0246()
L110B8:
    If  m001A = 1 Then Exit Sub
        Call sub045F("Keyword")
        l00CA% = extfn0B02("AOL Frame25", 0&)
            Do
            DoEvents
            If  m001A = 1 Then Exit Sub
            l00CE% = extfn0B10(l00CA%, "Keyword")
            Loop Until l00CE% > 0
        l00D0% = extfn0B24(l00CE%, "_AOL_Edit")
        l00D6 = extfn0AC9(l00D0%, mc00D2, 0, "aol://2719:2-2-" & List2.Text)
        l00D6 = extfn0AB5(l00D0%, mc00DA, 13, 0)
            Do
            DoEvents
            If  m001A = 1 Then Exit Sub
            l00DE% = extfn0B02("#32770", "America Online")
            Loop Until l00DE% > 0 Or fn0246() <> l00C6$
                If  l00DE% > 0 Then
                If  m001A = 1 Then Exit Sub
                l00DE% = extfn0B02("#32770", "America Online")
                l00D6 = extfn0AB5(l00DE%, mc00E0, 0, 0)
                GoTo L110B8
                ElseIf fn0246() <> l00C6$ Then
                If  m001A = 1 Then Exit Sub
                List2.RemoveItem List2.ListIndex
                m002C = 0
                sub060A .8#
                m002A% = Label3.Caption - Text3.Text
                GoTo L11418
                End If
        End If
End If

L11418:
If  Label3.Caption = "0" Then
sub055F "×–•[·FrostBite·]•–× The Mass Mail Has Begun!"
Command3D1.Enabled = True
Command3D2.Enabled = False
VBMsg1.SubClasshWnd = 0
MMMain.List1.Clear
For l00C0 = 0 To List1.ListCount - 1
MMMain.List1.AddItem List1.List(l00C0)
Next l00C0
MMMain.Frame3D1.Caption = "People " + MMMain.List1.ListCount
MMMain.Timer1.Enabled = True
Timer1.Enabled = False
Exit Sub
Unload Me
End If
sub055F "×–•[·FrostBite·]•–× FrostBite Mass Mailer v2.o"
sub060A .001#
sub055F "×–•[·FrostBite·]•–× Mass Mail Going Out in: " + Label3.Caption + ""
sub060A .001#
sub055F "×–•[·FrostBite·]•–× Say """ + Text2 + """ To Get On"
Label3.Caption = Val(Label3.Caption) - 1
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim l00F0 As String
Dim l00F2 As Integer
Dim l00F4 As String
Dim l00F6 As String
Dim l00F8 As Variant
l00F0$ = extfn0B38(lParam)
    l00F2% = InStr(1, l00F0$, ":")
    l00F4$ = Mid(l00F0$, 3, (l00F2% - 3))
    l00F6$ = Right(l00F0$, (Len(l00F0$) - Len(l00F4$)) - 4)
If  UCase$(l00F6$) = UCase$(Text2.Text) Then
For l00F8 = 0 To List1.ListCount - 1
If  UCase$(l00F4$) = UCase$(List1.List(l00F8)) Then Exit Sub
Next l00F8
List1.AddItem l00F4$
If  m002C = 0 Then
sub055F "×–•[·FrostBite·]•–× " + l00F4$ + " - Added to MM"
sub060A .6#
End If
Label4.Caption = "People: " + List1.ListCount
End If
End Sub

Sub VScroll1_Change ()
Label3.Caption = VScroll1.Value
End Sub
