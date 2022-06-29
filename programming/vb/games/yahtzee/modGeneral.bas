Attribute VB_Name = "modGeneral"
Option Explicit

'//Public Members
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'//Private Constants
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2

'//Public Constants
Global Const MyApp = "MTCYahtzee"  '//Registry key used for storing data

'//Public variables
Public Bonus63 As Integer   '//Left column bonus scoring countdown
Public lColTotal As Integer '//Left column total score
Public rColTotal As Integer '//Right column total score
Public GameTotal As Integer '//Game total
Public HSPosition As Integer '//High score position

'//Wav file playing subroutine
Public Sub PlaySound(strSound As String)
Dim wFlags%
    
  wFlags% = SND_ASYNC Or SND_NODEFAULT
  sndPlaySound strSound, wFlags%

End Sub

'//Get values from registry
Public Sub CheckReg()
Dim strCheckReg(2) As String
Dim midX As Long
Dim midY As Long

  '//Get center screen values
  midX = ((Screen.Width / 2) - (frmMain.Width / 2))
  midY = ((Screen.Height / 2) - (frmMain.Height / 2))
  
  '//Restore previous screen position
  frmMain.Left = GetSetting(MyApp, "Settings", "MainLeft", midX)
  frmMain.Top = GetSetting(MyApp, "Settings", "MainTop", midY)
  
  '//See if sound is on or off
  frmMain.mnuSound.Checked = GetSetting(MyApp, "Settings", "Sound", "True")
  
  '//If never played before add default values to registry
  strCheckReg(1) = GetSetting(MyApp, "Settings", "HSName1")
  strCheckReg(2) = GetSetting(MyApp, "Settings", "GamesPlayed")

  If strCheckReg(1) = "" And strCheckReg(2) = "" Then
    Call ResetHighScores
    Call ResetStats
  End If

End Sub

'//Add last game to statistics in registry
Public Sub UpdateStats(Score As Integer)
Dim Stat(2) As Long

  Stat(1) = CLng(GetSetting(MyApp, "Settings", "GamesPlayed", "0"))
  Stat(2) = CLng(GetSetting(MyApp, "Settings", "TotalScore", "0"))

  Stat(1) = CLng(Stat(1)) + 1
  Stat(2) = CLng(Stat(2)) + Score

  SaveSetting MyApp, "Settings", "GamesPlayed", Stat(1)
  SaveSetting MyApp, "Settings", "TotalScore", Stat(2)

End Sub

'//Reset statistics in registry
Public Sub ResetStats()
  
  SaveSetting MyApp, "Settings", "GamesPlayed", "0"
  SaveSetting MyApp, "Settings", "TotalScore", "0"

End Sub

'//Reset all high scores in registry!
Public Sub ResetHighScores()
  
  SaveSetting MyApp, "Settings", "HSName0", "Shannon Harmon"
  SaveSetting MyApp, "Settings", "HSName1", "Amy Kellar"
  SaveSetting MyApp, "Settings", "HSName2", "Thomas Allen"
  SaveSetting MyApp, "Settings", "HSName3", "John Baringer"
  SaveSetting MyApp, "Settings", "HSName4", "Denise Allen"
  SaveSetting MyApp, "Settings", "HSScore0", "300"
  SaveSetting MyApp, "Settings", "HSScore1", "275"
  SaveSetting MyApp, "Settings", "HSScore2", "250"
  SaveSetting MyApp, "Settings", "HSScore3", "225"
  SaveSetting MyApp, "Settings", "HSScore4", "200"

End Sub

'//See if last games total was high enough to be listed in top 5 players
Public Function CheckForHS() As Boolean
Dim Name0 As String, Name1 As String, Name2 As String, Name3 As String, Name4 As String
Dim Score0 As Integer, Score1 As Integer, Score2 As Integer, Score3 As Integer, Score4 As Integer
Dim i As Integer
Dim NewHSName As String

  Name0 = GetSetting(MyApp, "Settings", "HSName0", "Default")
  Name1 = GetSetting(MyApp, "Settings", "HSName1", "Default")
  Name2 = GetSetting(MyApp, "Settings", "HSName2", "Default")
  Name3 = GetSetting(MyApp, "Settings", "HSName3", "Default")
  Name4 = GetSetting(MyApp, "Settings", "HSName4", "Default")
  Score0 = GetSetting(MyApp, "Settings", "HSScore0", "Default")
  Score1 = GetSetting(MyApp, "Settings", "HSScore1", "Default")
  Score2 = GetSetting(MyApp, "Settings", "HSScore2", "Default")
  Score3 = GetSetting(MyApp, "Settings", "HSScore3", "Default")
  Score4 = GetSetting(MyApp, "Settings", "HSScore4", "Default")

  If GameTotal > Score4 Then
    
    NewHSName = InputBox("NEW HIGH SCORE" + Chr(13) + Chr(13) + "Please enter your name", "Yahtzee Deluxe")
    
    If GameTotal > Score4 And GameTotal <= Score3 Then
      
      SaveSetting MyApp, "Settings", "HSScore4", GameTotal
      SaveSetting MyApp, "Settings", "HSName4", NewHSName
      HSPosition = 4
    
    End If
    
    If GameTotal > Score3 And GameTotal <= Score2 Then
      
      SaveSetting MyApp, "Settings", "HSScore3", GameTotal
      SaveSetting MyApp, "Settings", "HSName3", NewHSName
      SaveSetting MyApp, "Settings", "HSScore4", Score3
      SaveSetting MyApp, "Settings", "HSName4", Name3
      HSPosition = 3
    
    End If
    
    If GameTotal > Score2 And GameTotal <= Score1 Then
      
      SaveSetting MyApp, "Settings", "HSScore2", GameTotal
      SaveSetting MyApp, "Settings", "HSName2", NewHSName
      SaveSetting MyApp, "Settings", "HSScore3", Score2
      SaveSetting MyApp, "Settings", "HSName3", Name2
      SaveSetting MyApp, "Settings", "HSScore4", Score3
      SaveSetting MyApp, "Settings", "HSName4", Name3
      HSPosition = 2
    
    End If
    
    If GameTotal > Score1 And GameTotal <= Score0 Then
      
      SaveSetting MyApp, "Settings", "HSScore1", GameTotal
      SaveSetting MyApp, "Settings", "HSName1", NewHSName
      SaveSetting MyApp, "Settings", "HSScore2", Score1
      SaveSetting MyApp, "Settings", "HSName2", Name1
      SaveSetting MyApp, "Settings", "HSScore3", Score2
      SaveSetting MyApp, "Settings", "HSName3", Name2
      SaveSetting MyApp, "Settings", "HSScore4", Score3
      SaveSetting MyApp, "Settings", "HSName4", Name3
      HSPosition = 1
    
    End If
    
    If GameTotal > Score0 Then
      
      SaveSetting MyApp, "Settings", "HSScore0", GameTotal
      SaveSetting MyApp, "Settings", "HSName0", NewHSName
      SaveSetting MyApp, "Settings", "HSScore1", Score0
      SaveSetting MyApp, "Settings", "HSName1", Name0
      SaveSetting MyApp, "Settings", "HSScore2", Score1
      SaveSetting MyApp, "Settings", "HSName2", Name1
      SaveSetting MyApp, "Settings", "HSScore3", Score2
      SaveSetting MyApp, "Settings", "HSName3", Name2
      SaveSetting MyApp, "Settings", "HSScore4", Score3
      SaveSetting MyApp, "Settings", "HSName4", Name3
      HSPosition = 0
    
    End If
    
    CheckForHS = True

  Else
  
    CheckForHS = False
  
  End If

End Function

