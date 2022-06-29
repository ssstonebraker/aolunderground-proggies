VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cXpos As Single
Private cYpos As Single
Private cXvel As Single
Private cYvel As Single
Private cYstart As Single
Private cYtilt As Single
Private cYtiltvel As Single
Private cYshadow As Single
Private cShape As Integer
Private cEnabled As Boolean

Property Get Xpos() As Single
Xpos = cXpos
End Property

Property Get Ypos() As Single
Ypos = cYpos
End Property

Property Get Xvel() As Single
Xvel = cXvel
End Property

Property Get Yvel() As Single
Yvel = cYvel
End Property

Property Get Ystart() As Single
Ystart = cYstart
End Property

Property Get Ytilt() As Single
Ytilt = cYtilt
End Property

Property Get Ytiltvel() As Single
Ytiltvel = cYtiltvel
End Property

Property Get Yshadow() As Single
Yshadow = cYshadow
End Property

Property Get Shape() As Integer
Shape = cShape
End Property

Property Get Enabled() As Boolean
Enabled = cEnabled
End Property

Property Let Xpos(X As Single)
cXpos = X
End Property

Property Let Ypos(X As Single)
cYpos = X
End Property

Property Let Xvel(X As Single)
cXvel = X
End Property

Property Let Yvel(X As Single)
cYvel = X
End Property

Property Let Ystart(X As Single)
cYstart = X
End Property

Property Let Ytilt(X As Single)
cYtilt = X
End Property

Property Let Ytiltvel(X As Single)
cYtiltvel = X
End Property

Property Let Yshadow(X As Single)
cYshadow = X
End Property

Property Let Shape(X As Integer)
cShape = X
End Property

Property Let Enabled(Y As Boolean)
cEnabled = Y
End Property
