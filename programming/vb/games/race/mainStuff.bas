Attribute VB_Name = "MainModule"
Global gbOkay As Byte 'the byte to test if the person has
'gone around the track or not

Global giFrameRate
Global Endloop As Boolean

Private starttime As Long, endtime As Long, mills As Long, ticks As Long, tickmod As Long, errorterm As Long, robomove As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function GetAsyncKeyState Lib "User32" (ByVal uAction As Long) As Long

' Our starting point.
' This sets up the control, and then loops until it
' is done




    
Sub main()
    ' Setup stuff
    SetupControl
    'Show Form
    Mainform.Show
    
    ' This is a function that continues to run
    ' until the program quits.  It pays attention
    ' to the form controls and does what it is told
    DoFrames
    
End Sub

Private Sub SetupControl()
    ' do setup
    'These are the defaults for the game
    Mainform!RenderAX1.PovActor 0
    Mainform!RenderAX1.FlatSet = App.Path + "\default.flt"
    Mainform!RenderAX1.BitmapSet = App.Path + "\default.bst"
    Mainform!RenderAX1.ActorDefs = App.Path + "\default.typ"
    Mainform!RenderAX1.TextureSet = App.Path + "\default.tdf"
    Mainform!RenderAX1.Level = App.Path + "\default.4dx" 'name of your level
    result = SelectPalette(Mainform.hdc, Mainform.RenderAX1.GetPalette, 0)
    result = RealizePalette(Mainform.hdc)
    BuildKeyboardMap
    giFrameRate = 0
End Sub

Private Sub DoFrames()
    ' it might seem easier to get the mouse buttons and movements
    ' through the mousedown and mousemove events, but that's not
    ' really the case. The RenderAX control doesn't handle mouse
    ' events, and Visual basic will only call mousedown and mousemove
    ' if the button is pressed outside the control. Hence, this
    ' roundabout way.
    endtime = dx_gettime
    
    
    Do While True
      result = SelectPalette(Mainform.hdc, Mainform.RenderAX1.GetPalette, 0)
      result = RealizePalette(Mainform.hdc)
      'Checks to see if a option has changed or not. This is so it doesn't slow down the
      'framerate
      
      
      
        'Renders here
        Mainform!RenderAX1.Render 0, 0
      
      DoEvents
      If Endloop = True Then
        Exit Do
      End If

      DoEvents
      'This is for the framerate
      giFrameRate = giFrameRate + 1
      
    Loop
   
End Sub


Function StuffToDo()
        'Checks the keyboard
        ScanKeyboard
        'Applys the keys pressed
        ApplyKeyboard
        'turns the player
        Mainform!RenderAX1.SpinActor 0, gdTurnAmount
        'moves the player
        Mainform!RenderAX1.MoveActor 0, gdCurrentSpeed
        'checks the player
        Mainform!RenderAX1.CheckActor 0
        
End Function
