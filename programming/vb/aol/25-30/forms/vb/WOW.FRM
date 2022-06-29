VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Wow - this must be that woodstock place!"
   ClientHeight    =   5820
   ClientLeft      =   1845
   ClientTop       =   2655
   ClientWidth     =   7365
   ClipControls    =   0   'False
   Height          =   6225
   Left            =   1785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   2310
   Width           =   7485
   WindowState     =   2  'Maximized
   Begin Timer Timer1 
      Interval        =   1
      Left            =   3075
      Top             =   2655
   End
End
Option Explicit

' Define arrays to hold the end points of the 2 lines
' 0 and 1 are for line 1, 2 and 3 are for line 2
Dim nXCoord(4) As Integer
Dim nYCoord(4) As Integer

Dim nXSpeed(4) As Integer
Dim nYSpeed(4) As Integer


' Define a variable to show how many trail lines should be shown
Dim nTrails As Integer

Sub Form_Click ()

    Unload frmMain

End Sub

Sub Form_Load ()

    Dim nIndex As Integer
    
    ' Initialise the points of the lines and their speeds
    
    For nIndex = 0 To 3

        nXCoord(nIndex) = frmMain.ScaleWidth \ 2
        nYCoord(nIndex) = frmMain.ScaleHeight \ 2

    Next

    ' Now set up the speeds, remember line 2 must be an exact copy of line 1
    nXSpeed(0) = -150: nXSpeed(2) = -150
    nXSpeed(1) = 70: nXSpeed(3) = 70

    nYSpeed(0) = -105: nYSpeed(2) = -105
    nYSpeed(1) = 90: nYSpeed(3) = 90


    ' Now set up the nTrails variable to show how many trail lines should be drawn
    nTrails = 50

End Sub

Sub Timer1_Timer ()

    Dim nIndex As Integer
    Dim nMaxIndex As Integer

    If nTrails > 0 Then
        nTrails = nTrails - 1
        nMaxIndex = 1
    Else
        nMaxIndex = 3
    End If


    For nIndex = 0 To nMaxIndex

        nXCoord(nIndex) = nXCoord(nIndex) + nXSpeed(nIndex)
        nYCoord(nIndex) = nYCoord(nIndex) + nYSpeed(nIndex)

        If nXCoord(nIndex) < 0 Or nXCoord(nIndex) > frmMain.ScaleWidth Then nXSpeed(nIndex) = -nXSpeed(nIndex)
        If nYCoord(nIndex) < 0 Or nYCoord(nIndex) > frmMain.ScaleHeight Then nYSpeed(nIndex) = -nYSpeed(nIndex)

    Next nIndex

    ' Now draw the lines, one in black, the other in blue
    Line (nXCoord(0), nYCoord(0))-(nXCoord(1), nYCoord(1)), &HFF0000
    Line (nXCoord(2), nYCoord(2))-(nXCoord(3), nYCoord(3)), &H0&

    DoEvents

End Sub

