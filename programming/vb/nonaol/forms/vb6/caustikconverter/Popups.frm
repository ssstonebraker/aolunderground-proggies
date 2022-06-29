VERSION 5.00
Begin VB.Form Popups 
   Caption         =   "caustik converter"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu men_presets 
      Caption         =   "Presets"
      Begin VB.Menu men_arialblack 
         Caption         =   "Arial"
         Begin VB.Menu arialblack_highres 
            Caption         =   "high resolution"
         End
      End
      Begin VB.Menu men_webdings 
         Caption         =   "WebDings"
         Begin VB.Menu webdings_squares 
            Caption         =   "squares (nice)"
         End
         Begin VB.Menu webdings_circles 
            Caption         =   "circles"
         End
      End
      Begin VB.Menu menwingdings 
         Caption         =   "WingDings"
         Begin VB.Menu wingsquares 
            Caption         =   "squares"
         End
      End
   End
   Begin VB.Menu men_about 
      Caption         =   "men_about"
      Begin VB.Menu abt 
         Caption         =   "About Caustik Converter"
      End
   End
End
Attribute VB_Name = "Popups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub abt_Click()
about_frm.aboutmovie.Base = App.Path
about_frm.aboutmovie.Movie = App.Path + "\about.swf"
about_frm.aboutmovie.BackgroundColor = RGB(0, 0, 0)
about_frm.aboutmovie.BGColor = RGB(0, 0, 0)
about_frm.Show (1)
about_frm.aboutmovie.GotoFrame (1)
about_frm.aboutmovie.Play
End Sub

Private Sub arialblack_highres_Click()
Call setOptions("|", "Arial", 80, 6)
End Sub



Private Sub webdings_circles_Click()
Call setOptions("n", "Webdings", 20, 1.5)
End Sub

Private Sub webdings_squares_Click()
Call setOptions("g", "Webdings", 20, 1.5)
End Sub
Private Function setOptions(lettery As String, fonty As String, resxy As Long, offsety)
On Error Resume Next
scan_letter = lettery
options_frm.fonty.text = lettery
DoEvents
scan_font = fonty
options_frm.Combo1.text = fonty
scan_resx = resxy
options_frm.xy = resxy
scan_offset = offsety
options_frm.yy = offsety
Call options_frm.Label18_MouseDown(1, 0, 0, 0)
End Function

Private Sub wingsquares_Click()
Call setOptions("Ì", "Wingdings", 18, 1.5)
End Sub
