Attribute VB_Name = "MainMod"
Public Declare Function PlaySndFx Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub Main() ' do splash routine
    Dim a(5) As String, pb As Integer
    Dim sh As Long, i As Integer
    sh = Screen.Height / 2: i = 250
    Dim n As Integer, k As Integer, z As Integer, p As Integer
    Splash.Show
    Splash.Refresh

    For n = 1 To sh Step i: k = k + i
    Splash.Top = n: Splash.Left = k
    Next n
    z = Splash.Label1.Top
    
    For p = 1 To 12: For n = 1 To 900: Splash.Label1.Top = n:
    Next n
    Next p: Splash.Label1.Top = z

    For n = sh To 1 Step -i: k = k + i
    Splash.Top = n: Splash.Left = k
    Next n
     Unload Splash ' end splash routine
    SR.Show ' turn program flow over to Form SR.frm
End Sub
    


