Attribute VB_Name = "MainGameMod"
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Module Holds API Call's, Variables and Constants For The Game '
'____________________________________________________________________'
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Global XSpeed As Integer, BallX As Integer
Global YSpeed As Integer, BallY As Integer
Global FastSpeed As Integer ' holds the fastest spped
Global NumBounces As Long   ' holds the total number of bounces
Global StartTime As Date
Global LivesLeft As Integer
Global GamePicsLoaded  As Boolean
Global TitlePicsLoaded  As Boolean
Global LoadPercent As Integer
Global ParentForm As Form
Global CmdSpeedParam As Integer
Public Const Clock = 1
Public Const AntiClock = 2
' Used to create shpes in the scrolling text
' to view these correctly change this font to "terminal"
Public Const B = "°", BB = "±", BBB = "²"
Public Const BBBB = "Û", RR = "ž", UpExcla = "þ"
Public Const LL = "­"
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Is The First Sub To Load Up   '
'____________________________________'
Public Sub Main()
Dim Result As VbMsgBoxResult
StopSounds True, True
' Get The Current Resolution
' and Load The Correct Form.
' I Decided to To Have One Small Form
' And One Large Form Rather The Changing The Res
' Because It Was Easier, And The Change Res Code
' gave me probs
'
' set the speed from the paramerter given from the launcher
'
If Trim(Command()) = "" Then
 MsgBox "   Please Start Quad-Ball From The Quad-Ball Launcher.   ", vbOKOnly, "Quad-Ball"
 Result = MsgBox("   Do You Want To Load The Quad-Ball Launcher ?   ", vbYesNo, "Quad-Ball")
 If Result = vbYes Then
   ThisDir
   Shell "LaunchQuadball.exe", vbNormalFocus
   End
  Else
   MsgBox "   Quad-Ball Will Now Exit   ", vbOKOnly, "Quad-Ball"
   End
 End If
Else
  CmdSpeedParam = Int(Val(Command()))
  If CmdSpeedParam < 5 Then CmdSpeedParam = 5
  'If CmdSpeedParam > 200 Then CmdSpeedParam = 200
End If
Dim MinTwipsX As Integer
Dim TotalTwipsX As Integer
MinTwipsX = Int(800 * Screen.TwipsPerPixelX)
TotalTwipsX = Screen.Width
If TotalTwipsX > MinTwipsX Then
 Set ParentForm = MainLarge
 Debug.Print ParentForm.caption
 Load MainLarge
ElseIf TotalTwipsX = MinTwipsX Then
 Set ParentForm = MainSmall
 Load MainSmall
ElseIf TotalTwipsX < MinTwipsX Then
 Dim NewLine As String
 NewLine = Chr(13) & Chr(13)
 MsgBox "This Game Requires A Resolution Of At Least 800 X 600." & NewLine & _
 "To Increase Your Resolution Follow These Steps:" & NewLine & _
 "1) Right Click On The Desktop." & NewLine & _
 "2) Select Properties From The Menu." & NewLine & _
 "3) Select The Settings Tab In The Dialog Which Appears." & NewLine & _
 "4) Slide The Screen Area Scroller To A Higher Resolution (i.e. 800 X 600)." & NewLine & _
 "5) If The ScrollBar Is Not There Your Monitor Doesn't Support The Reolution So You Can Not Play This Game.", _
  vbOKOnly, "Cannot Run, Contact Arvinder@Bigfoot.com For Further Help."
End
End If
End Sub
' Loads Scores From The Registry
Public Sub LoadScore()
 Static TopScore As String
 Static TopName As String
 TopScore = RegKeys.GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball", "TopScore")
 If Trim(TopScore) = "" Then
  ' if a score don't exist the make a "0" score
  TopScore = "0"
  Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\QuadBall", "TopScore", "0")
 End If
 ' show the sores in their captions
 ParentForm.HighestScore.caption = TopScore
 TopName = GetKeyValue(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball", "TopName")
 ParentForm.HighName.caption = TopName
End Sub

'Saves Scores To The Registry
Public Sub SaveScore(Name As String, Score As String)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball", "TopScore", Score)
 Call UpdateKey(HKEY_LOCAL_MACHINE, "software\ArviSehmi\Quadball", "TopName", Name)
End Sub
' Sub Is Used Loads All Pictures ( *.img )
Public Sub LoadPic(Destination As Object, File As String)
 On Error GoTo Handel1
 ' tell the loading bar to increase in percent
 LoadPercent = LoadPercent + 1
 LoadUp.caption = LoadPercent
 LoadUp.CurrLoad.caption = "Loading Pictures...( " & File & " )"
 LoadUp.Refresh
 Destination.Picture = LoadPicture(File)
 Exit Sub
Handel1:
MsgBox "Error:" & Chr(13) & Chr(13) & _
       "There Is a Missing File (" & File & _
       ") Which is Needed By This Game," & Chr(13) & _
       "Please Re-Install Quad-Ball, So The Error Can Be Corrected." & Chr(13) & Chr(13) & _
       " For Further Help Contact Arvinder@Bigfoot.Com", vbOKOnly, "Error, Missing File."
End
End Sub
' Sub Is Used to load Pictures ( *.img ) into a Picture Clips
Public Sub LoadAniPic(Destination As Object, SourceImg As PictureClip, Cell As Integer)
 On Error Resume Next
  ' tell the loading bar to increase in percent
 LoadPercent = LoadPercent + 1
 LoadUp.caption = LoadPercent
 LoadUp.Refresh
 LoadUp.CurrLoad.caption = "Loading Animated Pictures..."
 Destination.Picture = SourceImg.GraphicCell(Cell)
End Sub
Public Sub Highlight(Label As Label)
 If Label.Tag = "no" Then
 WAVPlay "click.qbs"
 Label.Tag = "yes"
 Label.Left = Label.Left - 10
 Label.FontSize = Label.FontSize + 5
 Label.ForeColor = RGB(0, 255, 0)
 End If
End Sub
Public Sub UnHighlight(Label As Label)
 If Label.Tag = "yes" Then
 Label.Tag = "no"
 Label.Left = Label.Left + 10
 Label.FontSize = Label.FontSize - 5
 Label.ForeColor = RGB(0, 90, 0)
 Else
 Label.Tag = "no"
 End If
End Sub
Public Sub Delay(TimeToPause As Single) ' Waits
 Dim TT As Double
 TT = Timer
 Do
  DoEvents
 Loop Until Timer > TT + TimeToPause
End Sub
Public Sub Sleep(TimeToPause As Single) ' Stops
 Dim TT As Double
 TT = Timer
 Do
 Loop Until Timer > TT + TimeToPause
End Sub

' Increase Ball Speed
' change the "SpeedToAdd" variable to two or three to make the game harder
' for pc's that are 200 mhz or less the speed should be set to 2, not 1
' if your pc is higher then 300 mhz then the speed should be set to 1 (estimated)
Public Sub IncSpeed(Optional SpeedToAdd As Integer = 1)
Dim XorY As Integer
Dim YSpeedTemp As Integer, XSpeedTemp As Integer
Randomize Timer
XorY = Int(Rnd * 2) ' gives a random value telling if X or Y should increase
If XorY = 0 Then
 If XSpeed > 0 Then XSpeed = XSpeed + SpeedToAdd Else XSpeed = XSpeed - SpeedToAdd
   XSpeedTemp = XSpeed ' inc X speed
  If XSpeedTemp > 0 Then XSpeedTemp = XSpeedTemp Else XSpeedTemp = -XSpeedTemp
 ElseIf XorY = 1 Then
  If YSpeed > 0 Then YSpeed = YSpeed + SpeedToAdd Else YSpeed = YSpeed - SpeedToAdd
   YSpeedTemp = YSpeed ' inc Y speed
  If YSpeedTemp > 0 Then YSpeedTemp = YSpeedTemp Else YSpeedTemp = -YSpeedTemp
End If
End Sub
'Set dir path to app's path
Public Sub ThisDir()
 ChDrive App.Path
 ChDir App.Path
End Sub
