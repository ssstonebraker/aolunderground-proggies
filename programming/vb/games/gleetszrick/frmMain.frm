VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zricks"
   ClientHeight    =   7200
   ClientLeft      =   1665
   ClientTop       =   735
   ClientWidth     =   9600
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then   ' Exit the game on ESC press
        If MsgBox("Are you sure you want to exit?", vbDefaultButton2 + vbYesNo + vbQuestion, "Zricks") = vbYes Then
            dixuAppEnd = True
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF12 Then  ' Exit the game on F12 press
        If MsgBox("Are you sure you want to exit?", vbDefaultButton2 + vbYesNo + vbQuestion, "Zricks") = vbYes Then
            dixuAppEnd = True
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF3 Then   ' Display the High Score Table
        Zricks_Pause
        frmScore.Show vbModal
        Zricks_Resume
    ElseIf KeyCode = vbKeyReturn Then  ' Check if the game is over.
        If InGameOver = True Then SetupNextLevel True
    ElseIf KeyCode = vbKeyP Then    ' Pause the game
        If Paused = False Then
            If InGameOver = False Then Zricks_Pause
        Else
            If InGameOver = False Then Zricks_Resume
        End If
    ElseIf KeyCode = vbKeyLeft Then  ' left arrow press
        If (bmpPaddle.X - Abs(bmpPaddle.VelocityX)) > 0 Then
            
            ' Discard any english the paddle might have had from the opposite direction
            If PaddleEnglish > 0 Then PaddleEnglish = 0
            
            PaddleEnglish = PaddleEnglish - 1
            
            bmpPaddle.VelocityX = -(MaxXSpeed - 1)
            bmpPaddle.Move
            bmpPaddle.Paint
        Else
            bmpPaddle.X = 0
            bmpPaddle.VelocityX = 0
            bmpPaddle.Paint
        End If
    ElseIf KeyCode = vbKeyRight Then ' right arrow press
        If (bmpPaddle.X + Abs(bmpPaddle.VelocityX) + paddleW) < Me.ScaleWidth Then
            
            If PaddleEnglish < 0 Then PaddleEnglish = 0
            
            PaddleEnglish = PaddleEnglish + 1
            
            bmpPaddle.VelocityX = Abs(MaxXSpeed - 1)
            bmpPaddle.Move
            bmpPaddle.Paint
        Else
            bmpPaddle.X = Me.ScaleWidth - paddleW
            bmpPaddle.VelocityX = 0
        End If
    ElseIf KeyCode = vbKeyS Then ' S key press, toggle the sound play flag.
        bSoundIn = Not bSoundIn
        If bSoundIn = True Then
            strSound = Chr(SC_SPACE) & Chr(SC_SPEAKER)
        Else
            strSound = ""
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then  ' left key up
    bmpPaddle.Paint
    bmpPaddle.VelocityX = 0
ElseIf KeyCode = vbKeyRight Then
    bmpPaddle.Paint
    bmpPaddle.VelocityX = 0
End If
End Sub

Private Sub Form_Load()
Unload frmSplash
Zricks_Init
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
dixuAppEnd = True
Zricks_Done
End Sub
