Attribute VB_Name = "KeyboardInfo"

'An AbZuTu circle.
Global Const CIRCLEMAX = 1023       'Number of divisions in a circle.
Global Const CIRCLESTEP = 0.3515625 '360/1024.
Global Const DegreesToAngle = 1024 / 360
Type abzu_circle
    x(1024) As Single
    y(1024) As Single
End Type
Global gtOffsetCircle As abzu_circle


'Vechile Moving stuff
Global Const Maxspeed = 50 'This is the maxspeed
Global Const Acceleration = 0.4 'This is how fast acceleration is
Global Const TurnSpeed = 0.75 'This is how fast turning is
Global Const MaxTurn = 7 'This is the max amount that you can be turning. In other words the wheel can't turn anymore
Global Const BrakeSpeed = 2 'How fast you can break
Global gdCurrentSpeed As Double 'The current speed
Global gdTurnAmount As Double 'The current turning amount basically how much the wheel is turned
Global Const MaxLook = 200 'How far the player can look up or down
'Define the keyboard map.
Type KeyBd
    MapLength As Integer
    KeyASCII(32) As Integer
    KeyFunction(32) As String
    KeyState(32) As String
    KeyDelay(32) As Integer
End Type
Global Const BIT15 = 32768          'For use with GetAsyncKeyState.
Global Const BIT1 = 1               'For use with GetAsyncKeyState.
Global Const PRESSEDDELAY = 10      'Delay before key is considered "PRESSED".
Global gtKeyboard As KeyBd
Declare Function GetAsyncKeyState Lib "User32" (ByVal uAction As Long) As Long
Function BuildKeyboardMap() As Integer
    'Randy Sanders
    'September 24, 1997.
    'The following structure and data map out the keyboard.
    'There are 14 keys mapped. You can add more,however, I would try to keep it under 32.
    
    
    'Accelerate.
    gtKeyboard.KeyDelay(1) = PRESSEDDELAY
    gtKeyboard.KeyASCII(1) = 38
    gtKeyboard.KeyFunction(1) = "ACCELERATE"
    gtKeyboard.KeyState(1) = "UP"
    
    'Brake.
    gtKeyboard.KeyDelay(2) = PRESSEDDELAY
    gtKeyboard.KeyASCII(2) = 40
    gtKeyboard.KeyFunction(2) = "BRAKE"
    gtKeyboard.KeyState(2) = "UP"
    
    'Turn left.
    gtKeyboard.KeyDelay(3) = PRESSEDDELAY
    gtKeyboard.KeyASCII(3) = 37
    gtKeyboard.KeyFunction(3) = "TURNLEFT"
    gtKeyboard.KeyState(3) = "UP"
    
    'Turn right.
    gtKeyboard.KeyDelay(4) = PRESSEDDELAY
    gtKeyboard.KeyASCII(4) = 39
    gtKeyboard.KeyFunction(4) = "TURNRIGHT"
    gtKeyboard.KeyState(4) = "UP"
    
        gtKeyboard.MapLength = 6
End Function
Function ScanKeyboard() As Integer
    Dim m%, n%
    
    'Randy Sanders
    'September 24, 1997
    'Here is where the keyboard gets scanned and states are assigned.
    'You can see how the number of keys mapped could affect preformance.
    
    For n = 0 To gtKeyboard.MapLength
        m = GetAsyncKeyState(gtKeyboard.KeyASCII(n))
        Select Case (m And BIT15)
            Case 0
                'This key must be up.
                gtKeyboard.KeyState(n) = "UP"
                gtKeyboard.KeyDelay(n) = PRESSEDDELAY
            Case Else
                'This key must be down.
                gtKeyboard.KeyDelay(n) = gtKeyboard.KeyDelay(n) - 1
                If gtKeyboard.KeyDelay(n) < 0 Then
                    gtKeyboard.KeyDelay(n) = 0
                    gtKeyboard.KeyState(n) = "PRESSED"
                Else
                    gtKeyboard.KeyState(n) = "DOWN"
                End If
        End Select
    Next
End Function

Function ApplyKeyboard() As Integer
    Dim m%, n%
    
    'Randy Sanders
    'September 24, 1997
    'This module supports three keyboard states.
    'Up, Down, and Pressed.
    'A "DOWN" keystate is automaticaly converted to a "PRESSED" state
    'after a certain amount of time determoned by the global constant "PRESSEDDELAY".
    For n = 0 To gtKeyboard.MapLength
        Select Case gtKeyboard.KeyState(n)
            Case "UP"
                m = KeyboardAction(gtKeyboard.KeyFunction(n), "UP")
            Case "PRESSED"
                m = KeyboardAction(gtKeyboard.KeyFunction(n), "PRESSED")
            Case "DOWN"
                m = KeyboardAction(gtKeyboard.KeyFunction(n), "DOWN")
        End Select
    Next
End Function
Function KeyboardAction(Action As String, State As String) As Integer
    Dim m%, n%
    Dim x As Single
    Dim y As Single
    Dim f As Single
    Dim Angle As Integer
    
    'Randy Sanders
    'September 24, 1997
    'Here is where keyboard states do their action.
    
    Select Case State
        Case "UP"
            'When the key is up, do this.
            Select Case Action
                
            End Select
        Case "DOWN"
            'When the key is down, do this.
            Select Case Action
                Case "ACCELERATE"
                    gdCurrentSpeed = gdCurrentSpeed + Acceleration
                    If gdCurrentSpeed > Maxspeed Then gdCurrentSpeed = Maxspeed
                    Mainform!lblSpeed.Caption = gdCurrentSpeed
                Case "BRAKE"
                    gdCurrentSpeed = gdCurrentSpeed - BrakeSpeed
                    If gdCurrentSpeed < 0 Then gdCurrentSpeed = 0
                    Mainform!lblSpeed.Caption = gdCurrentSpeed
                Case "TURNLEFT"
                    gdTurnAmount = gdTurnAmount - TurnSpeed
                    If gdTurnAmount < -1 * MaxTurn Then gdTurnAmount = -1 * MaxTurn
                Case "TURNRIGHT"
                    gdTurnAmount = gdTurnAmount + TurnSpeed
                    If gdTurnAmount > MaxTurn Then gdTurnAmount = MaxTurn
                
            End Select
    Case "PRESSED"
        'When the key is pressed, do this.
        Select Case Action
               Case "ACCELERATE"
                    gdCurrentSpeed = gdCurrentSpeed + Acceleration
                    If gdCurrentSpeed > Maxspeed Then gdCurrentSpeed = Maxspeed
                    Mainform!lblSpeed.Caption = gdCurrentSpeed
                Case "BRAKE"
                    gdCurrentSpeed = gdCurrentSpeed - BrakeSpeed
                    If gdCurrentSpeed < 0 Then gdCurrentSpeed = 0
                    Mainform!lblSpeed.Caption = gdCurrentSpeed
                Case "TURNLEFT"
                    gdTurnAmount = gdTurnAmount - TurnSpeed
                    If gdTurnAmount < -1 * MaxTurn Then gdTurnAmount = -1 * MaxTurn
                Case "TURNRIGHT"
                    gdTurnAmount = gdTurnAmount + TurnSpeed
                    If gdTurnAmount > MaxTurn Then gdTurnAmount = MaxTurn
                
            End Select
    End Select
End Function
