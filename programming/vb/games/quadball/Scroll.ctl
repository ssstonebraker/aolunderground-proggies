VERSION 5.00
Begin VB.UserControl ArviScroll 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   13.5
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   600
      Left            =   675
      TabIndex        =   0
      Top             =   1125
      Width           =   1725
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ArviScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Control Scrolls Text Across It '
'_____________________________________'
Dim ExitIt As Boolean
Dim Scrolling As Boolean
Dim sStr As String ' start string
' Calling This Start the Scroll
Public Sub StartScroll(Optional Speed As Single = 2.5)
 If Scrolling = True Then Exit Sub
 ExitIt = False
 sStr = Label1.caption
 Dim tStr As String    ' temporary string
 Dim NoLets As Integer ' number of letters
 Dim oStr As String    ' original string
 Dim i
 On Error Resume Next
 Do
 Scrolling = True
  For i = 0 To Speed * 1000
   DoEvents
   If ExitIt = True Then GoTo nd:
  Next i
  oStr = Label1.caption           '-----------
  NoLets = Len(Label1.caption)    '
  tStr = Left$(oStr, 1)            '  Get The Fisrt Letter
  oStr = Right$(oStr, (NoLets - 1)) ' And move it to the end.
  Label1 = oStr + tStr             '----------
  DoEvents
 Loop Until ExitIt = True
nd:
Scrolling = False
Label1 = sStr
Exit Sub
End Sub
Public Sub ExitScroll()
 ExitIt = True
 Scrolling = False
 Label1 = sStr
End Sub
Private Sub UserControl_Initialize()
 ExitIt = False
End Sub
Private Sub UserControl_Resize()
 Label1.Top = 0
 Label1.Left = 0 '-100
 Label1.Width = UserControl.Width '+ 200
 Label1.Height = UserControl.Height
End Sub

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' THE REST OF THE CODE IS USED FOR THE CONTROL'S '
' PROPERTIES, eg Text Colour, Font Size ...      '
'________________________________________________'
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Label1.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Label1.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Label1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get caption() As String
Attribute caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    caption = Label1.caption
End Property
Public Property Let caption(ByVal New_caption As String)
    Label1.caption() = New_caption
    PropertyChanged "caption"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Label1.Refresh
End Sub
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
    Label1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label1.caption = PropBag.ReadProperty("caption", "Label1")
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
End Sub
Private Sub UserControl_Terminate()
 ExitScroll
 ExitIt = True
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", Label1.BorderStyle, 0)
    Call PropBag.WriteProperty("caption", Label1.caption, "Label1")
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
End Sub

