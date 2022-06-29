VERSION 5.00
Begin VB.Form Rotations 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Faygo's Fun With Rotations"
   ClientHeight    =   5544
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6756
   DrawMode        =   6  'Mask Pen Not
   LinkTopic       =   "Form1"
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Txt2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1440
      Top             =   2760
   End
End
Attribute VB_Name = "Rotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Well here it is This example
'Was made by Faygo
'It will show you cool rotations
'to use with lines and points
'have fun
'-Faygo


Public Screen_X As Long
Public Screen_Y As Long
Public Half_X As Long
Public Half_Y As Long
Public AdZ As Integer

Private Type Point_3D
    X As Single
    Y As Single
    z As Single
End Type

Private Type Line_S
    Point1 As Integer
    Point2 As Integer
End Type
    
Dim P(8) As Point_3D
Dim l(12) As Line_S

Dim Cos_Tab(360) As Single, Sin_Tab(360) As Single

Private Function Rotate_Point(P As Point_3D, r1 As Integer, r2 As Integer, R3 As Integer) As Point_3D
    Dim Temp As Point_3D
    Dim T As Single
    Temp = P
    
    T = Temp.X
    
    Temp.X = T * Cos_Tab(r1) - Temp.Y * Sin_Tab(r1)
    Temp.Y = Temp.Y * Cos_Tab(r1) + T * Sin_Tab(r1)
    
    T = Temp.X
    Temp.X = T * Cos_Tab(r2) - Temp.z * Sin_Tab(r2)
    Temp.z = Temp.z * Cos_Tab(r2) + T * Sin_Tab(r2)

    T = Temp.z
    Temp.z = T * Cos_Tab(R3) - Temp.Y * Sin_Tab(R3)
    Temp.Y = Temp.Y * Cos_Tab(R3) + T * Sin_Tab(R3)
   
    Rotate_Point = Temp
End Function

Private Sub Line_3D(P1 As Point_3D, P2 As Point_3D)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim Col As Long
    
    Col = 128 - (P1.z - P2.z) * 4
    P1.z = P1.z + AdZ
    P2.z = P2.z + AdZ
    
    If (P1.z > 10 And P2.z > 10) Then
        X1 = P1.X * 256 / P1.z + Half_X
        X2 = P2.X * 256 / P2.z + Half_X
        Y1 = P1.Y * 256 / P1.z + Half_Y
        Y2 = P2.Y * 256 / P2.z + Half_Y
        Line (X1, Y1)-(X2, Y2), Col
    End If
    
    P1.z = P1.z - AdZ
    P2.z = P2.z - AdZ
End Sub

Private Sub Form_Load()
Txt1.Visible = True
Txt2.Visible = True
Txt3.Visible = True

Txt1.Text = "2"
Txt2.Text = "2"
Txt3.Text = "2"

    Dim Ctr As Integer
    For Ctr = 0 To 359
        Cos_Tab(Ctr) = Cos(Ctr * 3.1415 / 180)
        Sin_Tab(Ctr) = Sin(Ctr * 3.1415 / 180)
    Next Ctr

    Me.Show
       
    AdZ = 100
    
'    P(0).x = 0
'    P(0).y = 50
'    P(0).z = 0
'    P(1).x = 50
'    P(1).y = 0
'    P(1).z = 0
'    P(2).x = 0
'    P(2).y = 0
'    P(2).z = 50
'    P(3).x = -50
'    P(3).y = 0
'    P(3).z = 0
'    P(4).x = 0
'    P(4).y = 0
'    P(4).z = -50
'    P(5).x = 0
'    P(5).y = -50
'    P(5).z = 0
'    P(6).x = 10
'    P(6).y = 10
'    P(6).z = 10
'    P(7).x = -10
'    P(7).y = 10
'    P(7).z = 10
    
    
    P(0).X = -50
    P(0).Y = -50
    P(0).z = -50
    
    P(1).X = 50
    P(1).Y = -50
    P(1).z = -50
    
    P(2).X = 50
    P(2).Y = 50
    P(2).z = -50
    
    P(3).X = -50
    P(3).Y = 50
    P(3).z = -50
    
    P(4).X = -50
    P(4).Y = -50
    P(4).z = 50
    
    P(5).X = 50
    P(5).Y = -50
    P(5).z = 50
    
    P(6).X = 50
    P(6).Y = 50
    P(6).z = 50
    
    P(7).X = -50
    P(7).Y = 50
    P(7).z = 50
    
    l(0).Point1 = 0
    l(0).Point2 = 1
    l(1).Point1 = 1
    l(1).Point2 = 2
    l(2).Point1 = 2
    l(2).Point2 = 3
    l(3).Point1 = 4
    l(3).Point2 = 5
    l(4).Point1 = 5
    l(4).Point2 = 6
    l(5).Point1 = 6
    l(5).Point2 = 7
    
    l(6).Point1 = 0
    l(6).Point2 = 4
    l(7).Point1 = 1
    l(7).Point2 = 5
    l(8).Point1 = 2
    l(8).Point2 = 6
    l(9).Point1 = 3
    l(9).Point2 = 7
    
    l(10).Point1 = 0
    l(10).Point2 = 3
    l(11).Point1 = 4
    l(11).Point2 = 7

End Sub

Private Sub Form_Resize()
    Screen_X = Me.Width / Screen.TwipsPerPixelX
    Screen_Y = Me.Height / Screen.TwipsPerPixelY
    Half_X = Screen_X / 2
    Half_Y = Screen_Y / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
    Dim Ctr As Integer
    For Ctr = 0 To 7
        P(Ctr) = Rotate_Point(P(Ctr), Txt1.Text, Txt2.Text, Txt3.Text)
    Next Ctr
    Me.Cls
    For Ctr = 0 To 11
        Call Line_3D(P(l(Ctr).Point1), P(l(Ctr).Point2))
    Next Ctr
End Sub



