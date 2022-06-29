Attribute VB_Name = "basStopWatch"
Option Explicit
    Dim Seconds As Integer
    Dim Minutes As Integer
    Dim Hours As Integer

Public Sub StopWatch(TheLabel As Label)

    'This will run like a timer would on a stopwatch.
    'Just add your labels name in the parameter.
    'Example put this in a timer and set the timers_
    'interval to 1000 (or just a little less)
    'Call StopWatch(Label1)
    'Send Questions or comments to NightMare_36@hotmail.com
    'NightShade  04/25/1999
    

        Seconds% = Seconds% + 1
         
        If Seconds% >= 60 Then
            Minutes% = Minutes% + 1
            Seconds% = 0
        End If
        
        If Minutes% >= 60 Then
            Hours% = Hours% + 1
            Minutes = 0
        End If
         
        TheLabel.Caption = Format(Hours%, "00") & ":" & Format(Minutes%, "00") & ":" & Format(Seconds%, "00")
 

End Sub

