VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frogger in Visual Basic 6.0 - by Spliff"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrROAD 
      Interval        =   1
      Left            =   45
      Top             =   3330
   End
   Begin VB.Timer tmrWINRULES 
      Interval        =   1
      Left            =   585
      Top             =   4950
   End
   Begin VB.Timer tmrLOGS 
      Interval        =   1
      Left            =   180
      Top             =   585
   End
   Begin VB.Image imgFROG 
      Height          =   480
      Left            =   45
      Top             =   4905
      Width           =   480
   End
   Begin VB.Image imgBUS 
      Height          =   450
      Index           =   2
      Left            =   5445
      Top             =   3330
      Width           =   1125
   End
   Begin VB.Image imgBANG 
      Height          =   480
      Left            =   1080
      Top             =   4905
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLIMMO 
      Height          =   405
      Index           =   1
      Left            =   3600
      Top             =   4410
      Width           =   1800
   End
   Begin VB.Image imgLIMMO 
      Height          =   405
      Index           =   0
      Left            =   225
      Top             =   4410
      Width           =   1800
   End
   Begin VB.Image imgSPORTSCAR 
      Height          =   420
      Index           =   2
      Left            =   4770
      Top             =   3870
      Width           =   1440
   End
   Begin VB.Image imgSPORTSCAR 
      Height          =   420
      Index           =   1
      Left            =   2385
      Top             =   3870
      Width           =   1440
   End
   Begin VB.Image imgSPORTSCAR 
      Height          =   420
      Index           =   0
      Left            =   90
      Top             =   3870
      Width           =   1440
   End
   Begin VB.Image imgBUS 
      Height          =   450
      Index           =   1
      Left            =   3375
      Top             =   3285
      Width           =   1125
   End
   Begin VB.Image imgBUS 
      Height          =   450
      Index           =   0
      Left            =   1125
      Top             =   3285
      Width           =   1125
   End
   Begin VB.Label lblSTATUS 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "< MESSAGES >"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4155
      TabIndex        =   0
      Top             =   5130
      Width           =   1275
   End
   Begin VB.Image imgLOGS4 
      Height          =   420
      Index           =   2
      Left            =   3735
      Top             =   2205
      Width           =   1635
   End
   Begin VB.Image imgLOGS4 
      Height          =   420
      Index           =   3
      Left            =   5895
      Top             =   2205
      Width           =   1635
   End
   Begin VB.Image imgLOGS3 
      Height          =   420
      Index           =   3
      Left            =   4320
      Top             =   1665
      Width           =   1635
   End
   Begin VB.Image imgLOGS3 
      Height          =   420
      Index           =   2
      Left            =   2115
      Top             =   1665
      Width           =   1635
   End
   Begin VB.Image imgLOGS2 
      Height          =   420
      Index           =   2
      Left            =   3015
      Top             =   1125
      Width           =   1635
   End
   Begin VB.Image imgLOGS2 
      Height          =   420
      Index           =   3
      Left            =   5760
      Top             =   1125
      Width           =   1635
   End
   Begin VB.Image imgLOGS1 
      Height          =   420
      Index           =   2
      Left            =   2655
      Top             =   585
      Width           =   1635
   End
   Begin VB.Image imgLOGS1 
      Height          =   420
      Index           =   3
      Left            =   5400
      Top             =   585
      Width           =   1635
   End
   Begin VB.Image imgLOGS3 
      Height          =   420
      Index           =   1
      Left            =   45
      Top             =   1665
      Width           =   1635
   End
   Begin VB.Image imgLOGS2 
      Height          =   420
      Index           =   1
      Left            =   675
      Top             =   1125
      Width           =   1635
   End
   Begin VB.Image imgLOGS1 
      Height          =   420
      Index           =   1
      Left            =   45
      Top             =   585
      Width           =   1635
   End
   Begin VB.Image imgLOGS4 
      Height          =   420
      Index           =   1
      Left            =   945
      Top             =   2205
      Width           =   1635
   End
   Begin VB.Label roadlines 
      BackStyle       =   0  'Transparent
      Caption         =   "-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   4140
      Width           =   5505
   End
   Begin VB.Label roadlines 
      BackStyle       =   0  'Transparent
      Caption         =   "-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   3600
      Width           =   5505
   End
   Begin VB.Shape road 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   0
      Top             =   3240
      Width           =   5550
   End
   Begin VB.Image grass 
      Height          =   630
      Index           =   3
      Left            =   0
      Top             =   4860
      Width           =   5565
   End
   Begin VB.Image grass 
      Height          =   630
      Index           =   2
      Left            =   0
      Top             =   2700
      Width           =   5565
   End
   Begin VB.Image water 
      Height          =   2250
      Left            =   0
      Top             =   540
      Width           =   5610
   End
   Begin VB.Image grass 
      Height          =   630
      Index           =   1
      Left            =   0
      Top             =   -45
      Width           =   5565
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed1, speed2, speed3, speed4, speed5, speed6, speed7 As Integer
Dim frogmiddle As Integer

Private Sub Form_Load()
On Error GoTo something_is_wrong
    grass(1).Picture = frmLISTFILES.jpgGRASS.Picture 'LoadPicture(App.Path & "\" & "grass.jpg")
    grass(2).Picture = grass(1).Picture 'LoadPicture(App.Path & "\" & "grass.jpg")
    grass(3).Picture = grass(1).Picture 'LoadPicture(App.Path & "\" & "grass.jpg")
    
    water.Picture = frmLISTFILES.jpgWATER.Picture 'LoadPicture(App.Path & "\" & "water.jpg")
    imgLOGS1(1).Picture = frmLISTFILES.gifLOG.Picture  'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS1(2).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS1(3).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    
    imgLOGS2(1).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS2(2).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS2(3).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    
    imgLOGS3(1).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS3(2).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS3(3).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    
    imgLOGS4(1).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS4(2).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    imgLOGS4(3).Picture = imgLOGS1(1).Picture 'LoadPicture(App.Path & "\" & "log.gif")
    
    imgFROG.Picture = LoadPicture(App.Path & "\" & "frogg.ico")
    imgBANG.Picture = LoadPicture(App.Path & "\" & "dead.ico")
    
    imgBUS(0).Picture = frmLISTFILES.gifSCHOOLBUS.Picture 'LoadPicture(App.Path & "\" & "schoolbus.gif")
    imgBUS(1).Picture = imgBUS(0).Picture 'LoadPicture(App.Path & "\" & "schoolbus.gif")
    imgBUS(2).Picture = imgBUS(0).Picture 'LoadPicture(App.Path & "\" & "schoolbus.gif")
    
    imgSPORTSCAR(0).Picture = frmLISTFILES.gifSPORTSCAR.Picture 'LoadPicture(App.Path & "\" & "sportscar.gif")
    imgSPORTSCAR(1).Picture = imgSPORTSCAR(0).Picture 'LoadPicture(App.Path & "\" & "sportscar.gif")
    imgSPORTSCAR(2).Picture = imgSPORTSCAR(0).Picture 'LoadPicture(App.Path & "\" & "sportscar.gif")
    
    imgLIMMO(0).Picture = frmLISTFILES.gifLIMMO.Picture 'LoadPicture(App.Path & "\" & "limmo.gif")
    imgLIMMO(1).Picture = imgLIMMO(0).Picture 'LoadPicture(App.Path & "\" & "limmo.gif")

    imgFROG.Left = 45
    imgFROG.Top = 4905
    imgBANG.Visible = False
    tmrWINRULES.Enabled = True
    tmrLOGS.Enabled = True
    tmrROAD.Enabled = True
    Me.Height = 5790
    Me.Width = 5595
    Randomize
    speed1 = Int((100 * Rnd) + 30) ' Log speed
    speed2 = Int((100 * Rnd) + 30) ' Log speed
    speed3 = Int((100 * Rnd) + 30) ' Log speed
    speed4 = Int((100 * Rnd) + 30) ' Log speed
    speed5 = Int((80 * Rnd) + 30) ' Bus speed
    speed6 = Int((100 * Rnd) + 70) ' Sportcar speed
    speed7 = Int((90 * Rnd) + 60) ' Limmo speed
Exit Sub
    
something_is_wrong:
Dim err_reply As VbMsgBoxResult
    If Err.Number = 53 Then
        err_reply = MsgBox("One of the specified image files cannot be found.", vbRetryCancel + vbExclamation, "File not found")
        If err_reply = vbRetry Then Form_Load
        If err_reply = vbCancel Then End
    End If
    Exit Sub
End Sub

Private Sub Form_Keydown(KeyCode As Integer, shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: End 'Unload Me
        Case vbKeyLeft: imgFROG.Left = imgFROG.Left - 495
        Case vbKeyRight: imgFROG.Left = imgFROG.Left + 495
        Case vbKeyUp: imgFROG.Top = imgFROG.Top - 540
        Case vbKeyDown: imgFROG.Top = imgFROG.Top + 540
    End Select
    If imgFROG.Left <= grass(1).Left Then imgFROG.Left = imgFROG.Left + 495
    If (imgFROG.Left + imgFROG.Width) >= (grass(1).Left + grass(1).Width) Then imgFROG.Left = imgFROG.Left - 495
    If imgFROG.Top <= grass(1).Top Then imgFROG.Top = imgFROG.Top + 540
    If (imgFROG.Top + imgFROG.Height) >= (grass(1).Height + water.Height + grass(2).Height + grass(3).Height + road.Height) Then imgFROG.Top = imgFROG.Top - 540
    frogmiddle = imgFROG.Left + (imgFROG.Width / 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim closereply As VbMsgBoxResult
    tmrLOGS.Enabled = False
    tmrROAD.Enabled = False
    closereply = MsgBox("Are you sure you wish to exit?", vbQuestion + vbYesNo, "Exit?")
    If closereply = vbYes Then End
    If closereply = vbNo Then
        Cancel = 1
        Form_Load
    End If
End Sub

Private Sub tmrLOGS_Timer()
    imgLOGS1(1).Left = imgLOGS1(1).Left + speed1
    imgLOGS1(2).Left = imgLOGS1(2).Left + speed1
    imgLOGS1(3).Left = imgLOGS1(3).Left + speed1
    imgLOGS2(1).Left = imgLOGS2(1).Left - speed2
    imgLOGS2(2).Left = imgLOGS2(2).Left - speed2
    imgLOGS2(3).Left = imgLOGS2(3).Left - speed2
    imgLOGS3(1).Left = imgLOGS3(1).Left + speed3
    imgLOGS3(2).Left = imgLOGS3(2).Left + speed3
    imgLOGS3(3).Left = imgLOGS3(3).Left + speed3
    imgLOGS4(1).Left = imgLOGS4(1).Left - speed4
    imgLOGS4(2).Left = imgLOGS4(2).Left - speed4
    imgLOGS4(3).Left = imgLOGS4(3).Left - speed4
    If imgLOGS1(1).Left >= Me.Width Then imgLOGS1(1).Left = 0 - imgLOGS1(1).Width
    If imgLOGS1(2).Left >= Me.Width Then imgLOGS1(2).Left = 0 - imgLOGS1(2).Width
    If imgLOGS1(3).Left >= Me.Width Then imgLOGS1(3).Left = 0 - imgLOGS1(3).Width
    If imgLOGS2(1).Left <= 0 - imgLOGS2(1).Width Then imgLOGS2(1).Left = Me.Width
    If imgLOGS2(2).Left <= 0 - imgLOGS2(2).Width Then imgLOGS2(2).Left = Me.Width
    If imgLOGS2(3).Left <= 0 - imgLOGS2(3).Width Then imgLOGS2(3).Left = Me.Width
    If imgLOGS3(1).Left >= Me.Width Then imgLOGS3(1).Left = 0 - imgLOGS1(1).Width
    If imgLOGS3(2).Left >= Me.Width Then imgLOGS3(2).Left = 0 - imgLOGS1(2).Width
    If imgLOGS3(3).Left >= Me.Width Then imgLOGS3(3).Left = 0 - imgLOGS1(3).Width
    If imgLOGS4(1).Left <= 0 - imgLOGS4(1).Width Then imgLOGS4(1).Left = Me.Width
    If imgLOGS4(2).Left <= 0 - imgLOGS4(2).Width Then imgLOGS4(2).Left = Me.Width
    If imgLOGS4(3).Left <= 0 - imgLOGS4(3).Width Then imgLOGS4(3).Left = Me.Width
    If imgFROG.Top = 585 Then imgFROG.Left = imgFROG.Left + speed1
    If imgFROG.Top = 1125 Then imgFROG.Left = imgFROG.Left - speed2
    If imgFROG.Top = 1665 Then imgFROG.Left = imgFROG.Left + speed3
    If imgFROG.Top = 2205 Then imgFROG.Left = imgFROG.Left - speed4
End Sub

Private Sub tmrROAD_Timer() 'Bus, sportcar and limmo movement
    imgBUS(0).Left = imgBUS(0).Left + speed5
    imgBUS(1).Left = imgBUS(1).Left + speed5
    imgBUS(2).Left = imgBUS(2).Left + speed5
    imgSPORTSCAR(0).Left = imgSPORTSCAR(0).Left - speed6
    imgSPORTSCAR(1).Left = imgSPORTSCAR(1).Left - speed6
    imgSPORTSCAR(2).Left = imgSPORTSCAR(2).Left - speed6
    imgLIMMO(0).Left = imgLIMMO(0).Left + speed7
    imgLIMMO(1).Left = imgLIMMO(1).Left + speed7
    
    If imgBUS(0).Left >= Me.Width Then imgBUS(0).Left = 0 - imgBUS(0).Width
    If imgBUS(1).Left >= Me.Width Then imgBUS(1).Left = 0 - imgBUS(1).Width
    If imgBUS(2).Left >= Me.Width Then imgBUS(2).Left = 0 - imgBUS(2).Width
    If imgSPORTSCAR(0).Left <= 0 - imgSPORTSCAR(0).Width Then imgSPORTSCAR(0).Left = Me.Width
    If imgSPORTSCAR(1).Left <= 0 - imgSPORTSCAR(1).Width Then imgSPORTSCAR(1).Left = Me.Width
    If imgSPORTSCAR(2).Left <= 0 - imgSPORTSCAR(2).Width Then imgSPORTSCAR(2).Left = Me.Width
    If imgLIMMO(0).Left >= Me.Width Then imgLIMMO(0).Left = 0 - imgLIMMO(0).Width
    If imgLIMMO(1).Left >= Me.Width Then imgLIMMO(1).Left = 0 - imgLIMMO(1).Width
End Sub

Private Sub tmrWINRULES_Timer()
Dim winreply As VbMsgBoxResult
    If imgFROG.Top = 45 Then 'Finishing grass (END)
        lblSTATUS.Caption = "Well Done! You saved the frog."
        tmrLOGS.Enabled = False
        tmrROAD.Enabled = False
        winreply = MsgBox("Well done player. Would you like to play again?", vbYesNo + vbQuestion, "Complete!")
        If winreply = vbYes Then
            lblSTATUS.Caption = "Loading..."
            Form_Load
        End If
        If winreply = vbNo Then End
    End If

    If imgFROG.Top = 585 Then 'Row of logs #1
        If frogmiddle < imgLOGS1(3).Left And frogmiddle > (imgLOGS1(2).Left + imgLOGS1(2).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS1(2).Left And frogmiddle > (imgLOGS1(1).Left + imgLOGS1(1).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS1(1).Left And frogmiddle > (imgLOGS1(3).Left + imgLOGS1(3).Width) Then GoTo youfucked
        'imgFROG.Left = imgFROG.Left + speed1
        frogmiddle = imgFROG.Left + (imgFROG.Width / 2)
    End If
    
    If imgFROG.Top = 1125 Then 'Row of logs #2
        If frogmiddle < imgLOGS2(3).Left And frogmiddle > (imgLOGS2(2).Left + imgLOGS2(2).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS2(2).Left And frogmiddle > (imgLOGS2(1).Left + imgLOGS2(1).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS2(1).Left And frogmiddle > (imgLOGS2(3).Left + imgLOGS2(3).Width) Then GoTo youfucked
        'imgFROG.Left = imgFROG.Left - speed2
        frogmiddle = imgFROG.Left + (imgFROG.Width / 2)

    End If
    
    If imgFROG.Top = 1665 Then 'Row of logs #3
        If frogmiddle < imgLOGS3(3).Left And frogmiddle > (imgLOGS3(2).Left + imgLOGS3(2).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS3(2).Left And frogmiddle > (imgLOGS3(1).Left + imgLOGS3(1).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS3(1).Left And frogmiddle > (imgLOGS3(3).Left + imgLOGS3(3).Width) Then GoTo youfucked
        'imgFROG.Left = imgFROG.Left + speed3
        frogmiddle = imgFROG.Left + (imgFROG.Width / 2)

    End If
    
    If imgFROG.Top = 2205 Then 'Row of logs #4
        If frogmiddle < imgLOGS4(3).Left And frogmiddle > (imgLOGS4(2).Left + imgLOGS4(2).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS4(2).Left And frogmiddle > (imgLOGS4(1).Left + imgLOGS4(1).Width) Then GoTo youfucked
        If frogmiddle < imgLOGS4(1).Left And frogmiddle > (imgLOGS4(3).Left + imgLOGS4(3).Width) Then GoTo youfucked
        'imgFROG.Left = imgFROG.Left - speed4
        frogmiddle = imgFROG.Left + (imgFROG.Width / 2)

    End If
    
    If imgFROG.Top = 2745 Then 'Grass before water (NEARLY THERE)
        lblSTATUS.Caption = "Keep going!!!"
    End If
    
    If imgFROG.Top = 3285 Then 'Bus lane of road (TOP)
        If frogmiddle > imgBUS(0).Left And frogmiddle < (imgBUS(0).Left + imgBUS(0).Width) Then GoTo youfucked
        If frogmiddle > imgBUS(1).Left And frogmiddle < (imgBUS(1).Left + imgBUS(1).Width) Then GoTo youfucked
        If frogmiddle > imgBUS(2).Left And frogmiddle < (imgBUS(2).Left + imgBUS(2).Width) Then GoTo youfucked
    End If
    
    If imgFROG.Top = 3825 Then 'Sportscar lane of road (MIDDLE)
        If frogmiddle > imgSPORTSCAR(0).Left And frogmiddle < imgSPORTSCAR(0).Left + imgSPORTSCAR(0).Width Then GoTo youfucked
        If frogmiddle > imgSPORTSCAR(1).Left And frogmiddle < imgSPORTSCAR(1).Left + imgSPORTSCAR(1).Width Then GoTo youfucked
        If frogmiddle > imgSPORTSCAR(2).Left And frogmiddle < imgSPORTSCAR(2).Left + imgSPORTSCAR(2).Width Then GoTo youfucked
    End If
    
    If imgFROG.Top = 4365 Then 'Limmon lane of road (BOTTOM)
        If frogmiddle > imgLIMMO(0).Left And frogmiddle < imgLIMMO(0).Left + imgLIMMO(0).Width Then GoTo youfucked
        If frogmiddle > imgLIMMO(1).Left And frogmiddle < imgLIMMO(1).Left + imgLIMMO(1).Width Then GoTo youfucked
    End If
    
    If imgFROG.Top = 4905 Then 'Grass before road (START)
        lblSTATUS.Caption = "Go for it!!!"
    End If
    Exit Sub

youfucked:
Dim losereply As VbMsgBoxResult
    imgBANG.Left = imgFROG.Left
    imgBANG.Top = imgFROG.Top
    imgBANG.Visible = True
    lblSTATUS.Caption = "Oh dear, you died!"
    tmrLOGS.Enabled = False
    tmrROAD.Enabled = False
    losereply = MsgBox("Oh dear, you died! Would you like to play again?", vbYesNo + vbQuestion, "Ooops!")
    If losereply = vbYes Then
        lblSTATUS.Caption = "Loading..."
        Form_Load
    End If
    If losereply = vbNo Then End
Exit Sub
End Sub
