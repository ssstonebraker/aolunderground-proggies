VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "BeRsUrK Mp3 pLaYeR bY Renegade"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   4920
   End
   Begin VB.PictureBox ProgressBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      DrawWidth       =   40
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   80
      Left            =   720
      ScaleHeight     =   45
      ScaleWidth      =   3585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Mp3 Files Loaded"
      Top             =   2760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8454143
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Path."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Count."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Name."
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Played"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Visit Renegade's World!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      MousePointer    =   10  'Up Arrow
      TabIndex        =   16
      ToolTipText     =   "Visit Renegade's World!"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   8
      Left            =   1800
      Top             =   2040
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   15
      Left            =   4080
      Picture         =   "frmMain.frx":014A
      Stretch         =   -1  'True
      ToolTipText     =   "Normal Play"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   14
      Left            =   1920
      Picture         =   "frmMain.frx":06B4
      ToolTipText     =   "Clear Playlist"
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label lblPlayMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Normal play"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Play Mode"
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label lblPlayMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Play Mode."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Width           =   795
   End
   Begin VB.Image imgControl 
      Height          =   150
      Index           =   13
      Left            =   4320
      Picture         =   "frmMain.frx":0D0E
      Top             =   2400
      Width           =   180
   End
   Begin VB.Image imgControl 
      Height          =   150
      Index           =   12
      Left            =   570
      Picture         =   "frmMain.frx":11F8
      Top             =   2400
      Width           =   180
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   5
      Left            =   5040
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   4
      Left            =   5040
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   3
      Left            =   5040
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   2
      Left            =   5040
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   1
      Left            =   5040
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape VolumeInd 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   0
      Left            =   5040
      Top             =   2400
      Width           =   135
   End
   Begin VB.Image imgControl 
      Height          =   375
      Index           =   11
      Left            =   4800
      Picture         =   "frmMain.frx":16E2
      ToolTipText     =   "Decrease Volume"
      Top             =   2160
      Width           =   180
   End
   Begin VB.Image imgControl 
      Height          =   375
      Index           =   10
      Left            =   4800
      Picture         =   "frmMain.frx":1CBC
      ToolTipText     =   "Increase Volume"
      Top             =   1800
      Width           =   180
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   7
      Left            =   1080
      Top             =   2040
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   9
      Left            =   1200
      Picture         =   "frmMain.frx":2296
      ToolTipText     =   "Random Select"
      Top             =   2040
      Width           =   510
   End
   Begin VB.Shape speCommandLight 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   6
      Left            =   360
      Top             =   2040
      Width           =   105
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   8
      Left            =   480
      Picture         =   "frmMain.frx":28F0
      ToolTipText     =   "Mute Track"
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label lblArtistA 
      BackColor       =   &H00400000&
      Caption         =   "No Track loaded."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      ToolTipText     =   "Mp3 Playing"
      Top             =   240
      Width           =   3360
   End
   Begin VB.Label lblPositionA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      ToolTipText     =   "Position of Mp3"
      Top             =   720
      Width           =   90
   End
   Begin VB.Label lblLengthA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Length of selected Mp3"
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   ":"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "How long song has been playing"
      Top             =   960
      Width           =   45
   End
   Begin VB.Label lblELTimeD 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "How long song has been playing"
      Top             =   960
      Width           =   90
   End
   Begin VB.Label lblELTimeC 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      ToolTipText     =   "How long song has been playing"
      Top             =   960
      Width           =   90
   End
   Begin VB.Label lblELTimeB 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      ToolTipText     =   "How long song has been playing"
      Top             =   960
      Width           =   90
   End
   Begin VB.Label lblELTimeA 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "How long song has been playing"
      Top             =   960
      Width           =   90
   End
   Begin VB.Label lblElTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Elapsed Time."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00400000&
      Caption         =   "Position."
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00400000&
      Caption         =   "Artist."
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblLength 
      BackColor       =   &H00400000&
      Caption         =   "Length."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   555
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   720
      X2              =   4370
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   720
      X2              =   4370
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   7
      Left            =   4320
      Picture         =   "frmMain.frx":2F4A
      ToolTipText     =   "Fast Play"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   6
      Left            =   3840
      Picture         =   "frmMain.frx":34B4
      Stretch         =   -1  'True
      ToolTipText     =   "Slow Play"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   480
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   480
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape speBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   975
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":3A1E
      ToolTipText     =   "Previous track"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   1
      Left            =   1200
      Picture         =   "frmMain.frx":4078
      ToolTipText     =   "Play Mp3"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   2
      Left            =   1920
      Picture         =   "frmMain.frx":46D2
      ToolTipText     =   "Pause Track"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   3
      Left            =   2640
      Picture         =   "frmMain.frx":4D2C
      ToolTipText     =   "Stop Track"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   4
      Left            =   3360
      Picture         =   "frmMain.frx":5386
      ToolTipText     =   "Next Track"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgControl 
      Height          =   180
      Index           =   5
      Left            =   4080
      Picture         =   "frmMain.frx":59E0
      ToolTipText     =   "Open Mp3(s)"
      Top             =   1800
      Width           =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   240
      X2              =   5310
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   6135
      Y2              =   2760
   End
   Begin VB.Shape InfoWindow 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape speBorder 
      BorderColor     =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   1
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sReturnBuffer As String * 30



'#####################################################################################
'# yo wuzz up if u dled da source u readin diz lol diz wont work for nuttin but VB6  #
'# so u bes have it or letz jes say                                                  #
'# it 'loads as a .bas'                                                              #
'# I kinda like the i-face I made...it took me about 3 months to figure out          #
'# howto make one of these.  I made sure it looks complicated, but really, its not.  #
'# you just hafto know your standard MCI functions                                   #
'# I like the coding because it's not like other mp3 players                         #
'# I think it makes it run smoother...I dunno                                        #
'# The reason being I'm not directing you on this is because really I wanna be nice  #
'# and mean at the same time itz cool  I letcha have the source now YOU figure it    #
'# out :) cuz u gotta be kiddin me if im gunna direct ya im lazy! :-P                #
'#####################################################################################
Private Sub Form_Load()

    Dim lCreateRegion As Long
        lCreateRegion& = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 30, 30)

    SetWindowRgn Me.hwnd, lCreateRegion&, True


    mVariables.iVolumeSetting = 498

    VolumeInd(0).FillColor = RGB(250, 0, 0)
    VolumeInd(1).FillColor = RGB(250, 0, 0)
    VolumeInd(2).FillColor = RGB(250, 0, 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If mVariables.bTrackIsPlaying = True Then
        Call imgControl_Click(3)
    End If

End Sub

Private Sub Form_Resize()
        
    On Error Resume Next
    If Me.Width <> 5670 Then
        Me.Width = 5670
    ElseIf Me.Height <> 6900 Then
        Me.Height = 6900
    End If
    
End Sub

Private Sub imgControl_Click(Index As Integer)

    Dim lShortPath As Long
    Dim sShortPath As String * 260
    Dim sShortPathName As String
    
        
        Select Case Index

            Case 0
                With lstFiles
                    If .ListItems.Count > 1 Then
                        If .SelectedItem.Index > 1 Then
                            
                            .ListItems(.SelectedItem.Index - 1).Selected = True
                        
                            Call imgControl_Click(3)
                            Call imgControl_Click(1)
                        
                        End If
                    End If
                End With
                                                                                    
            Case 1
                If lstFiles.ListItems.Count > 0 Then
                    Me.MousePointer = 11
                    ProgressBar.Cls
                    
                    With lstFiles
                        lShortPath& = GetShortPathName(.ListItems(.SelectedItem.Index).Text & .ListItems(.SelectedItem.Index).ListSubItems(2).Text, sShortPath$, 260)
                        sShortPathName$ = mProcFunc.ftnStripNullChar(sShortPath$)
                    End With

                    mciSendString "open " & sShortPathName$ & " type MPEGVideo alias mp3", 0, 0, 0
                    mciSendString "set mp3 time format tmsf", 0, 0, 0
                    mciSendString "play mp3", 0, 0, 0
                    mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
                    mciSendString "setaudio mp3 volume to " & mVariables.iVolumeSetting, 0, 0, 0
                    lblLengthA.Caption = sReturnBuffer$
                    mVariables.lTrackLength = Val(sReturnBuffer)
                    With lstFiles
                        lblArtistA.Caption = "(" & Mid(.ListItems(.SelectedItem.Index).ListSubItems(1).Text, 1, 2) & ") " & .ListItems(.SelectedItem.Index).ListSubItems(2).Text
                        .ListItems(.SelectedItem.Index).ListSubItems(3).Text = "P"
                    End With
                    mVariables.bTrackIsPlaying = True
                    
                    Timer1.Enabled = True
                    Me.MousePointer = 0
                    
                End If
                
            Case 2
                If lstFiles.ListItems.Count > 0 Then
                    mciSendString "pause mp3", 0, 0, 0
                    Timer1.Enabled = False
                    mVariables.bTrackIsPlaying = False
                End If
                    
            Case 3
                Timer1.Enabled = False
                mciSendString "stop mp3", 0, 0, 0
                mciSendString "close all", 0, 0, 0
                lblLengthA.Caption = "0"
                lblPositionA.Caption = "0"
                lblELTimeA.Caption = "0"
                lblELTimeB.Caption = "0"
                lblELTimeC.Caption = "0"
                lblELTimeD.Caption = "0"
                lblArtistA.Caption = "No Track loaded."
                ProgressBar.Cls
                mVariables.bTrackIsPlaying = False
            Case 4
                If lstFiles.ListItems.Count > 0 Then
                    If lstFiles.SelectedItem.Index < lstFiles.ListItems.Count Then
                        
                        lstFiles.ListItems(lstFiles.SelectedItem.Index + 1).Selected = True
                        
                        Call imgControl_Click(3)
                        Call imgControl_Click(1)
                    
                    End If
                End If
                        
            Case 5
                Me.MousePointer = 11
                frmExplore.Show (1)
                Me.MousePointer = 0
                
            Case 6
                mciSendString "set mp3 speed 500", 0, 0, 0
                lblPlayMode(1).Caption = "Slow play"
            
            Case 7
                mciSendString "set mp3 speed 1500", 0, 0, 0
                lblPlayMode(1).Caption = "Fast play"
        
            Case 8
                If mVariables.bTrackIsPlaying = True Then
                    If mVariables.bAudioAllOff = False Then
                        mciSendString "set mp3 audio all off", 0, 0, 0
                        mVariables.bAudioAllOff = True
                    Else
                        mciSendString "set mp3 audio all on", 0, 0, 0
                        mVariables.bAudioAllOff = False
                    End If
                End If
                    
            Case 9
            If mVariables.bRandomSet = True Then
                mVariables.bRandomSet = False
            Else
                mVariables.bRandomSet = True
            End If
            
            Case 10
                
                mProcFunc.subSetVolume ("Increase")
            
            Case 11
            
                mProcFunc.subSetVolume ("Decrease")
            
            Case 12
            
            Case 13
            
            Case 14
                lstFiles.ListItems.Clear
                
            Case 15
                mciSendString "set mp3 speed 1000", 0, 0, 0
                lblPlayMode(1).Caption = "Normal play"
        
        End Select


End Sub

Private Sub imgControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
            
    Select Case Index
            
        Case 0
                    
        Case 1
        
        Case 2
        
        Case 3
        
        Case 4
        
        Case 5
        
        Case 6
        
        Case 7
                
        Case 8
            
        Case 9
        
        Case 10
        
        Case 11
        
        Case 12
        
        Case 13
        
        Case 14
            
        Case 15
    
    End Select

End Sub

Private Sub imgControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    

    
    Select Case Index
                                                                                    
        Case 0
            If lstFiles.ListItems.Count > 1 Then
                mVariables.byCommandLight = 0
            End If
                        
        Case 1
            If lstFiles.ListItems.Count > 0 Then
                mVariables.byCommandLight = 1
            End If
            
        Case 2
            If mVariables.bTrackIsPlaying = True Then
                mVariables.byCommandLight = 2
            End If
        
        Case 3
            If mVariables.bTrackIsPlaying = True Then
                mVariables.byCommandLight = 3
            End If
        
        Case 4
            If lstFiles.ListItems.Count > 1 Then
                mVariables.byCommandLight = 4
            End If
    
        Case 5

            mVariables.byCommandLight = 5
        
        Case 6
        
        Case 7
        
        Case 8
            If mVariables.bTrackIsPlaying = True Then
                If mVariables.bAudioAllOff = True Then
                Else
                End If
            End If
            mVariables.byCommandLight = 6
        
        Case 9
            If mVariables.bRandomSet = True Then
            Else
            End If
            mVariables.byCommandLight = 7
            
        Case 10
        
        Case 11
        
        Case 12
        
        Case 13
        
        Case 14
            mVariables.byCommandLight = 8
            
        Case 15
    
    End Select

End Sub

Private Sub Label2_Click()
Shell ("start http://renegadesworld.cjb.net")
End Sub

Private Sub lstFiles_DblClick()
    
    Call imgControl_Click(3)
    Call imgControl_Click(1)
    
End Sub


Private Sub ProgressBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim lPosition As Long

    lPosition& = Val(lblLengthA.Caption) / ProgressBar.Width * x

    mciSendString "play mp3 from " & lPosition&, 0, 0, 0

    lPosition& = 0
        
    lstFiles.SetFocus

End Sub

Private Sub subPicScan(lMaxValue As Long, sPercent As Single)
    
    On Error Resume Next
    
    With ProgressBar
        .Cls
        .DrawMode = 13
        .CurrentX = .Width / 2 - .TextWidth("   ") / 2
        .CurrentY = .Height - 255
        .DrawMode = 10
        ProgressBar.Line (-200, 30)-Step(.Width * sPercent \ lMaxValue + 200, 0), RGB(0, 0, 110), BF
        .Refresh
    End With

End Sub


Private Sub Timer1_Timer()
    
    mciSendString "status mp3 position", sReturnBuffer$, Len(sReturnBuffer$), 0
    lblPositionA.Caption = sReturnBuffer$
    subPicScan (mVariables.lTrackLength), (Val(sReturnBuffer$))
    lblELTimeD.Caption = Val(lblELTimeD.Caption) + 1
    If Val(lblELTimeD.Caption) > Val(9) Then
        lblELTimeC.Caption = Val(lblELTimeC.Caption) + 1
        lblELTimeD.Caption = "0"
    End If
    If lblELTimeC.Caption > 5 Then
        lblELTimeB.Caption = Val(lblELTimeB.Caption) + 1
        lblELTimeC.Caption = "0"
    End If
    If Val(lblPositionA.Caption) >= Val(lblLengthA.Caption) Then
        mVariables.bTrackIsPlaying = False
        lblPositionA.Caption = "0"
        lblLengthA.Caption = "0"
                
        With lstFiles
            
            If .ListItems.Count > 1 And .SelectedItem.Index < .ListItems.Count And mVariables.bRandomSet = False Then
                
                .ListItems(.SelectedItem.Index + 1).Selected = True
                
                Call imgControl_Click(3)
                
                Call imgControl_Click(1)
            
            Else
                Call imgControl_Click(3)
            End If
            
            If mVariables.bRandomSet = True Then

                Dim iRandomSelect As Integer

                iRandomSelect% = mProcFunc.ftnRandomSelect

                .ListItems(iRandomSelect%).Selected = True
                
                Call imgControl_Click(3)
                Call imgControl_Click(1)

            End If
            
        End With
        
    End If

End Sub
