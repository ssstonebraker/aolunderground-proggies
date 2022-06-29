VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Time Project"
   ClientHeight    =   5340
   ClientLeft      =   720
   ClientTop       =   1320
   ClientWidth     =   4335
   Height          =   5775
   Icon            =   CHP22.FRX:0000
   Left            =   645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   4335
   Top             =   960
   Width           =   4485
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   3240
      TabIndex        =   11
      Top             =   1560
      Width           =   852
   End
   Begin CommandButton cmndSet 
      Caption         =   "&Set System"
      Height          =   372
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   1332
   End
   Begin Timer timeMouse 
      Interval        =   2500
      Left            =   3360
      Top             =   0
   End
   Begin PictureBox pictTheWindsOfTime 
      AutoRedraw      =   -1  'True
      Height          =   1332
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   3825
      TabIndex        =   7
      Top             =   3840
      Width           =   3852
   End
   Begin Timer timeClock 
      Interval        =   250
      Left            =   3000
      Top             =   0
   End
   Begin CommandButton cmndCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1092
   End
   Begin TextBox textThen 
      Height          =   408
      Left            =   240
      TabIndex        =   1
      Text            =   "textThen"
      Top             =   1080
      Width           =   3852
   End
   Begin Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Now (Serial)"
      Height          =   252
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   1212
   End
   Begin Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Now (String)"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1212
   End
   Begin Label lablMrMouse 
      BackStyle       =   0  'Transparent
      Caption         =   "lablMrMouse"
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3732
   End
   Begin Label lablStopWatch 
      BackColor       =   &H00000000&
      Caption         =   "lablStopWatch"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   1092
   End
   Begin Label lablSerial 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   492
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   1092
   End
   Begin Label lablDifference 
      BackStyle       =   0  'Transparent
      Caption         =   "lablDifference"
      Height          =   732
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3732
   End
   Begin Label lablNow 
      BackStyle       =   0  'Transparent
      Caption         =   "lablNow"
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1212
   End
End

Sub cmndCalculate_Click ()
     On Error GoTo BadDate                        'set the trap
     thenDate = CVDate(textThen.Text)             'convert date the user typed in
     On Error GoTo 0                              'success!
     yr = Str$(Year(thenDate) - Year(Now))        'difference in years
     mt = Str$(Month(thenDate) - Month(Now))      'difference in months
     dy = Str$(Day(thenDate) - Day(Now))          'difference in days
     hr = Str$(Hour(thenDate) - Hour(Now))        'difference in hours
     mn = Str$(Minute(thenDate) - Minute(Now))    'difference in minutes
     sc = Str$(Second(thenDate) - Second(Now))    'difference in seconds
     wd = Weekday(thenDate)                       'get the new day of the week, then choose the right string
     dow = Choose(wd, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

     msg = yr & " years, " & mt & " months, " & dy & " days; " & Chr$(10) & " "     'put together the label
     msg = msg & hr & " hours, " & mn & " minutes, " & sc & " seconds " & Chr$(10)
     msg = msg & "The day of the week you gave is a " & dow
     lablDifference.Caption = msg
Exit Sub

BadDate:
     Beep
     textThen.SetFocus
Exit Sub

End Sub

Sub cmndExit_Click ()
     End                 'Ends the program
End Sub

Sub cmndSet_Click ()
     Static oldTime As Double, newTime As Double    'these allow us to reset the time
     Select Case cmndSet.Caption                    'are we setting or resetting?
          Case "&Set System"                        'set the system clock
               oldTime = Now                        'mark the start of the new time
               On Error GoTo BadSetTime             'set trap
               setTime = CVDate(textThen.Text)      'new date and time
               Date = setTime                       'set the date
               Time = setTime                       'set the time
               On Error GoTo 0                      'successful!
               cmndSet.Caption = "&Reset System"    'now indicate need to reset
          Case "&Reset System"                      'reset time back to where we were
               setTime = oldTime + Now - newTime    'difference between old and new
               Date = setTime                       'set date
               Time = setTime                       'set time
               cmndSet.Caption = "&Set System"      'reset the caption
     End Select
     newTime = Now                                  'the new time after being set
Exit Sub

BadSetTime:
     Beep
     textThen.SetFocus
Exit Sub

End Sub

Sub Form_Load ()
     textThen.Text = "Type a date and time in me!"
     lablDifference.Caption = "I'll display the difference between the system date and the date and time you type above."
     lablMrMouse.Caption = "I display the mouse position as a date (x) and time (y)"
     pictTheWindsOfTime.Scale (0, 0)-(1, 1)            'for mouse movement
     pictTheWindsOfTime.Print
     pictTheWindsOfTime.Print "   Click Me!"
     timeClock.Tag = Str$(Timer)                          'remember where we start (for stopwatch)
End Sub

Sub Form_Unload (Cancel As Integer)
     If cmndSet.Caption = "&Reset System" Then    'if we still need to reset the time,
          cmndSet_Click                           'do it before unloading
     End If
End Sub

Sub pictTheWindsOfTime_Click ()
     timeMouse.Enabled = True      'stop updating lablMrMouse for a couple of seconds
End Sub

Sub pictTheWindsOfTime_MouseMove (Button As Integer, Shift As Integer, x As Single, y As Single)
     If timeMouse.Enabled = False Then            'if user hasn't clicked, update lablMrMouse's display
          yr = 1994                               'keep it to 1994
          mt = Int(x * 12)                        'horizontal mouse position represents the date
          dy = Int(x * 365)
          hr = Int(y * 24)                        'vertical mouse position represents the time
          mn = Int(y * 60 * 24) Mod 60
          sc = Int(y * 60 * 24 * 24) Mod 60
          mouseTime = DateSerial(yr, mt, dy) + TimeSerial(hr, mn, sc)  'put together the date and time
          lablMrMouse.Caption = Format$(mouseTime, "dddd, mmmm d, yyyy   h:mm:ss am/pm") ' and display it
     End If
End Sub

Sub textThen_GotFocus ()
     textThen.SelStart = 0                   'select all the text
     textThen.SelLength = Len(textThen.Text)
End Sub

Sub timeClock_Timer ()
     lablNow.Caption = Date$ & "  " & Time$            'generic functions for right now
     dt = DateValue(Now)                               'extract date serial number
     tm = TimeValue(Now)                               'extract time serial number
     lablSerial = Format$(dt, "00,0000") & Format$(tm, "   .0000000000") 'format as serial numbers, not as date variants!
     startTime = timeClock.Tag                            'this is when program started
     elapsed = Timer - startTime                       'how much time has gone by?
     elapsed = (elapsed \ 60) * 100 + elapsed Mod 60   'time is measured in base 60
     lablStopWatch = Format$(elapsed, "00:00")         'and update the stopwatch
End Sub

Sub timeMouse_Timer ()
     timeMouse.Enabled = False     'start updating lablMrMouse again
End Sub

