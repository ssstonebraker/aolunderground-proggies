VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   MaskColor       =   &H00FFFFFF&
   MouseIcon       =   "Button.ctx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ToolboxBitmap   =   "Button.ctx":030A
   Begin VB.PictureBox Clicked 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   3
      Top             =   2430
      Width           =   1050
   End
   Begin VB.PictureBox Offit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   75
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   2
      Top             =   1845
      Width           =   1050
   End
   Begin VB.PictureBox Show 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      MouseIcon       =   "Button.ctx":061C
      MousePointer    =   99  'Custom
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   0
      Width           =   1275
   End
   Begin VB.PictureBox Onover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   75
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      Top             =   1275
      Width           =   1050
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Ok people.....The reason why I'm
'giving out the coding to this OCX
'is because...I want everyone to learn
'How to use BitBlt....That is a great function!
'It can be used for taking screen shoots...Making
'Pictures Transparent, copying the picture in one
'picture box to another (QUICKLY!). I also wanted
'people to know how to make an OCX or "Custom control"
'on there own...This was my first and I had a hard time
'getting help from people...So I asked on AOL's vb message board
'(Which by the way helps out alot more than the "vb" chat room)
'and I got my help really quick well except for the SetCapture,
'GetCapture and ReleaseCapture....I got some other help for that
'Ne how....Use this AND LEARN FROM IT!!!! PLEASE LEARN FROM IT!
'If you do take this coding and add to it...Please keep my name
'to this though....I put some effort into this and wish to have
'some credit for it!...I'm going to make a version 2.0 of this
'OCX soon...But its not going to have ALOT more options :-( sorry
'If you have ne questions about this "Ctl" file I'm going to try
'and get an e-mail address for programming help only (The e-mail address
'well be posted on my web page as soon as I make it...k? k!)
'Oh yeah....Last of all...this was first put on my freind Iris's web page
'at www.come.to/iris_n_aka....then I put it on my page at www.move.to/LafreakPlace
'"I'm evil, not cruel"~Lafreak
Option Explicit
Public Event Click() 'this makes the "Click" event
Public Event DDlClick() 'makes the Double click event
Public Event MouseOVER(Button As Integer, Shift As Integer, X As Single, Y As Single) 'makes the Mouse over event
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'makes the mouse down event
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'makes the Mouse up event
Private Const MouseON = "Pic_over" 'Makes "MouseON" a private const (basicly like a value that can't be changed)
Private Const mouseoff = "Pic_off" 'Makes "mouseoff" a private const (basicly like a value that can't be changed)
Private Const MouseCLICK = "Pic_down" 'Makes "mouseclick"  a private const (basicly like a value that can't be changed)
'

Public Property Get Pic_over() As Picture
Set Pic_over = Onover.Picture 'makes the property "Pic_Over" = whats ever in Onover
End Property

Public Property Set Pic_over(Picv As Picture)
UserControl.Onover.Picture = Picv 'sets Onover to the value passed to "Pic_over" (Picv)
PropertyChanged MouseON 'lets the project know that "MouseON" has changed
UserControl.Height = Show.ScaleHeight * 15 'makes the control the same height as Show's height
UserControl.Width = Show.ScaleWidth * 15 'makes the control the same width as Show's height
End Property

Public Property Get Pic_off() As Picture
Set Pic_off = Offit.Picture 'sets the property "Pic_off" to equal whats ever in offit
End Property

Public Sub AfterClick()
'the point behind this sub is to call it
'after a click or something where the mouse
'isn't on the picture and its up to you
'to tell this Control the mouse isn't on it
'Its used like this....example
'Button1.AfterClick
'Thats it :-) then The picture is back to
'the picture you selected for property
'"Pic_off"
Show.Picture = Offit.Picture 'makes Show equal Offit
End Sub
Public Property Get Pic_down() As Picture
Set Pic_down = Clicked.Picture 'makes property Pic_down equal what ever clicked equals
End Property

Public Property Set Pic_off(Picv As Picture)
UserControl.Offit.Picture = Picv 'makes Offit = Picv (the value passed to this property by the user)
UserControl.Show.Picture = Picv 'makes Show = Picv (the value passed to this property by the user)
PropertyChanged mouseoff 'Lets the project know that "Mouseoff" has changed
UserControl.Height = Show.ScaleHeight * 15 'makes the control's height the same as Show's height
UserControl.Width = Show.ScaleWidth * 15 'makes the control's width the same as show's width
End Property

Public Property Set Pic_down(Picv As Picture)
UserControl.Clicked.Picture = Picv 'makes Clicked = Picv (the value passed to this property by the user)
PropertyChanged MouseCLICK 'Lets the project know that "MouseCLICK" has changed
UserControl.Height = Show.ScaleHeight * 15 'makes the control's height the same as show's height
UserControl.Width = Show.ScaleWidth * 15 'makes the control's width the same as show's width
End Property

Private Sub Clicked_Click()

End Sub

Private Sub Onover_Click()

End Sub

Private Sub show_Click()
'this is how the Events are made...
'Now what its doing is making it so
'when ever the picture show is clicked
'it raise's the event "Click" for the
'user...Dunno if that makes sence to ya :-/
RaiseEvent Click 'raises the event "Click" for the user
End Sub

Private Sub Show_DblClick()
'this is how the Events are made...
'Now what its doing is making it so
'when ever the picture show is clicked
'it raise's the event "Ddlclick" for the
'user...Dunno if that makes sence to ya :-/
RaiseEvent DDlClick 'raises the event "DDlClick" for the user
End Sub

Private Sub Show_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is where it all gets a little tricky....
X = BitBlt(Show.hDC, 0, 0, Clicked.ScaleWidth, Clicked.ScaleHeight, Clicked.hDC, 0, 0, SRCCOPY)
'ok that line above makes Show equal whats in Clicked
'this is the exacts about how it works....
'#1. "Show.hDC" tells the function "BitBlt" where we want to copy the picture to
'#2. "0" thats the X position of where the picture is going to be placed in "Show"
'#3. "0" thats the y position of where the picture is going to be placed in "Show"
'#4. "Clicked.ScaleWidth" tells the function BitBlt the width of the picture to copy to "Show"
'#5. "Clicked.ScaleHeight" tells the function BitBlt the width of the picture to copy to "Show"
'#6. "Clicked.hDC" tells the function BitBlt where its copying the picture from
'#7. "0" tells the BitBlt function the x position of where to copy from "Clicked"
'#8. '#7. "0" tells the BitBlt function the y position of where to copy from "Clicked"
'#9. (Ghal where finnally there!) tells the BitBlt function what method of copying where using...In this case "SRCCOPY" which is copying
Show.Refresh
'refresh Show (So the new image is there!)
RaiseEvent MouseDown(Button, Shift, X, Y) 'Raise's the event "MouseDown"
'this is how the Events are made...
'Now what its doing is making it so
'when ever the picture show is clicked
'it raise's the event "MouseDown" for the
'user...Dunno if that makes sence to ya :-/
End Sub

Private Sub Show_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'WWWWWOOOOOOOHHHHHH IT GETS HARD IN HERE! I needed to search for a little bit of help for this part!
Dim Ret As Long 'lets the project know value Ret is a "Long" value
    If GetCapture() <> Show.hwnd Then 'if Getcapture ISN'T show.hwnd THEN
        Ret = SetCapture(Show.hwnd) 'Set capture of Show
        Call BitBlt(Show.hDC, 0, 0, Onover.ScaleWidth, Onover.ScaleHeight, Onover.hDC, 0, 0, SRCCOPY)
        'Ghal I don't want to write about this again! Look back in the "MouseDown" event
        Show.Refresh
        'Refresh show (that way the picture will show!)
    End If
    'Ends the if
    If X >= 0 And X <= Show.ScaleWidth And Y >= 0 And Y <= Show.ScaleHeight Then 'Ok all that just said "If the mouse is in the picture then"
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = Show.hwnd Then 'If capture is "Show" then
            Ret = ReleaseCapture() 'Release the capture
            Call BitBlt(Show.hDC, 0, 0, Offit.ScaleWidth, Offit.ScaleHeight, Offit.hDC, 0, 0, SRCCOPY)
            'Ghal I don't want to write about this again! Look back in the "MouseDown" event
            Show.Refresh
            'Refresh "show" (that way the picture will show!)
        End If 'end the if we made
    End If 'end the if we made
    'this is how the Events are made...
'Now what its doing is making it so
'when ever the picture show is clicked
'it raise's the event "MouseOver" for the
'user...Dunno if that makes sence to ya :-/
RaiseEvent MouseOVER(Button, Shift, X, Y) 'Raises the event "MouseOver"

End Sub

Private Sub Show_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = BitBlt(Show.hDC, 0, 0, Onover.ScaleWidth, Onover.ScaleHeight, Onover.hDC, 0, 0, SRCCOPY)
'I said it twice and I'll say it again....I don't want to write about this function again so look in the event "MouseDown"
Show.Refresh
'Refreshes "Show" (That way the image will show)
'this is how the Events are made...
'Now what its doing is making it so
'when ever the picture show is clicked
'it raise's the event "Click" for the
'user...Dunno if that makes sence to ya :-/
RaiseEvent MouseUp(Button, Shift, X, Y) 'Raises the event MouseUp
End Sub

Private Sub UserControl_InitProperties()
Set Pic_over = Picture 'Makes property Pic_over = Picture (This isn't too badly needed)
Set Pic_off = Picture 'Makes property Pic_off = Picture (This isn't too badly needed)
Set Pic_down = Picture 'Makes property Pic_down = Picture (This isn't too badly needed)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = BitBlt(UserControl.Show.hDC, 0, 0, UserControl.Offit.Width, UserControl.Offit.Height, UserControl.Offit.hDC, 0, 0, SRCCOPY)
'Check the event "MouseDown" in show for an explanation of the function "BitBlt"
Show.Refresh
'Refreshes Show (makes it so the image will show)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set Pic_down = PropBag.ReadProperty(MouseCLICK, Picture) 'makes the property Pic_down = What ever the user sent to "MouseCLICK"/Pic_down
Set Pic_over = PropBag.ReadProperty(MouseON, Picture) 'makes the property Pic_over = What ever the user sent to "MouseON"/Pic_over
Set Pic_off = PropBag.ReadProperty(mouseoff, Picture) 'makes the property Pic_off = What ever the user sent to "MouseOff"/Pic_off
Call UserControl_Resize 'calls for the control to be resized (according to the coding in "Usercontrol_Resize")
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Show.ScaleHeight * 15 'makes the control the same height as "show"'s height
UserControl.Width = Show.ScaleWidth * 15 'makes the control the same width as "show"'s Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag 'makes it so I don't have to type "Propbag." and whatever all the time
.WriteProperty MouseCLICK, Pic_down 'makes the property "Pic_down" avalible
.WriteProperty MouseON, Pic_over 'makes the property "Pic_over" avalible
.WriteProperty mouseoff, Pic_off 'makes the property "Pic_off" avalible
End With 'ends the with so I have to type "Propbag." and whatever again
End Sub
