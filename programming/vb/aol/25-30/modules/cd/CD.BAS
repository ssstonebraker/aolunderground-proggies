' CD.FRM
Option Explicit

Sub Command1_Click ()
Dim l0026 As Variant
Dim l002A As Variant
Dim l0040 As Variant
    MousePointer = 11
    On Error GoTo LBFA
    MMControl1.Command = "Open"
    On Error GoTo 0
    l0026 = False
    MMControl1.TimeFormat = l002A
    Command1.Enabled = False
    Command2.Enabled = True
    Label2.Caption = Str$(MMControl1.Tracks)
    If  Label2.Caption = 1 Then
        l0040 = "One track, may not be an Audio CD!"
        MsgBox l0040, 64, "CD Player"
    End If
    MousePointer = 0
    Exit Sub

LBFA:
    subC8
    Resume LC10

LC10:
    Unload CD
End Sub

Sub Command2_Click ()
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
    l0048 = False
    l004C = MMControl1.Tracks
    Select Case Val(l004C)
    Case 1
        l0050 = "One track, may not be an Audio CD!"
        MsgBox l0050, 64, "Shuffle CD Player Example"
        Call MMControl1_StopClick(-1)
    Case 2 To 36
        Call sub370
    Case Is > 36
        l0050 = "Only support for 36 tracks!  Unable to Shuffle!"
        MsgBox l0050, 64, "Shuffle CD Player Example"
        Call MMControl1_StopClick(-1)
    End Select
End Sub

Sub Command3_Click ()
End
End Sub

Sub Form_Load ()
Dim l0056 As Integer
Dim l005C As String
   l0056% = extfn3E8(CD.hWnd, -1, 0, 0, 0, 0, 3)
    If  App.PrevInstance Then
        l005C$ = "Shuffle CD Player Example already running!"
        MsgBox l005C$, 48
        End
    End If
    MMControl1.Wait = True
    MMControl1.UpdateInterval = 0
    Command2.Enabled = False
    Label2.Caption = "0"
    MMControl1.DeviceType = "CDAudio"
End Sub

Sub Form_Unload (p1 As Integer)
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    End
End Sub

Sub MMControl1_EjectClick (Cancel As Integer)
Dim l0062 As Variant
    l0062 = False
    Command1.Enabled = True
    Command2.Enabled = False
    Label2.Caption = "0"
    MMControl1.UpdateInterval = 0
    On Error GoTo LF5E
    MMControl1.Command = "Eject"
    MMControl1.Command = "Close"
    On Error GoTo 0
    Exit Sub

LF5E:
    subC8
    Resume Next
End Sub

Sub MMControl1_PauseClick (Cancel As Integer)
Dim l0068 As Variant
    l0068 = False
    MMControl1.UpdateInterval = 0
End Sub

Sub MMControl1_PlayClick (Cancel As Integer)
Dim l006E As Variant
    l006E = False
    MMControl1.UpdateInterval = 1000
End Sub

Sub MMControl1_StatusUpdate ()
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
    l0072 = MMControl1.Position And &HFF&
    l0076 = (MMControl1.Position And &HFF00&) \ &H100
    l007A = (MMControl1.Position And &HFF0000&) \ &H10000&
    l007E = (MMControl1.Position And &H7F000000&) \ &H1000000&
    If  (MMControl1.Position And &H80000000&) <> 0 Then
        l007E = l007E + &H80
    End If
    l0082 = l0072
    MMControl1.Track = l0082
    l0086 = MMControl1.TrackLength And &HFF&
    l008A = (MMControl1.TrackLength And &HFF00&) \ &H100
    If  (MMControl1.TrackLength And &H80000000&) <> 0 Then
        l008A = l008A + &H80
    End If
    If  MMControl1.Mode = l008E Then
        If  l0092 = True Then
        Else
            Call MMControl1_StopClick(-1)
        End If
    End If
End Sub

Sub MMControl1_StopClick (Cancel As Integer)
Dim l0098 As Variant
Dim l009C As Variant
    l0098 = False
    MMControl1.UpdateInterval = 0
    MMControl1.To = MMControl1.Start
    MMControl1.Command = "Seek"
    MMControl1.Track = 1
    l009C = 1
End Sub

Sub sub370 ()
Dim l00A0 As Variant
Dim l00A4 As Integer
Dim l00A6 As Integer
Dim l00A8 As Integer
Dim l00AA As Integer
Dim l00AC() As Integer
Dim l00B2() As Integer
    l00A0 = Val(Label2.Caption)
    ReDim l00AC(1 To l00A0)        As Integer
    ReDim l00B2(1 To l00A0)     As Integer
    Randomize
    For l00A6 = 1 To l00A0
        l00AC(l00A6) = l00A6
    Next l00A6
    l00A4 = l00A0
    For l00A6 = 1 To l00A4
        l00AA = Int(l00A4 * Rnd + 1)
        l00B2(l00A6) = l00AC(l00AA)
        For l00A8 = l00AA To l00A4 - 1
            l00AC(l00A8) = l00AC(l00A8 + 1)
        Next l00A8
        l00A4 = l00A4 - 1
    Next l00A6
    For l00A6 = 1 To l00A0
        Debug.Print l00B2(l00A6); " ";
    Next l00A6
    Debug.Print
End Sub
