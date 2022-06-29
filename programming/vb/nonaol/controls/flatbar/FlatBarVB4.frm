VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   4155
   ClientTop       =   2355
   ClientWidth     =   6690
   Height          =   4260
   Left            =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   Top             =   2010
   Width           =   6810
   Begin FlatBar32.FlatBar FlatBar1 
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   609
      ButnCount       =   13
      BtnStyle0       =   1
      ImageNumber1    =   1
      BtnStyle1       =   1
      ImageNumber2    =   2
      BtnStyle2       =   1
      ImageNumber3    =   3
      BtnStyle3       =   1
      ImageNumber4    =   4
      BtnStyle4       =   1
      ImageNumber5    =   5
      BtnStyle5       =   1
      ImageNumber6    =   6
      BtnStyle6       =   1
      ImageNumber7    =   7
      BtnStyle7       =   1
      ImageNumber8    =   8
      BtnStyle8       =   1
      ImageNumber9    =   9
      BtnStyle9       =   1
      ImageNumber10   =   10
      BtnStyle10      =   1
      ImageNumber11   =   11
      BtnStyle11      =   1
      ImageNumber12   =   12
      BtnStyle12      =   1
      ImageNumber13   =   13
      BtnStyle13      =   1
      Picture         =   "FlatBarVB4.frx":0000
      MaskColor       =   16776960
      Wrappable       =   -1  'True
   End
   Begin FlatBar32.FlatBar FlatBar2 
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   609
      ButnCount       =   7
      BtnStyle0       =   1
      ImageNumber1    =   1
      BtnStyle1       =   1
      ImageNumber2    =   2
      BtnStyle2       =   1
      ImageNumber3    =   3
      BtnStyle3       =   1
      ImageNumber4    =   4
      BtnStyle4       =   1
      ImageNumber5    =   5
      BtnStyle5       =   1
      ImageNumber6    =   6
      BtnStyle6       =   1
      ImageNumber7    =   7
      BtnStyle7       =   1
      Picture         =   "FlatBarVB4.frx":1712
      MaskColor       =   16776960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
'Needed for runtime
FlatBar1.RunTime
FlatBar2.RunTime
'
'Place in form Resize
'FlatBar1.ResizeRebar me

FlatBar1.CreateRebar Me, True
FlatBar1.AddBandsToRebar ChildWindow:=FlatBar1.GetToolbarHwnd, BandText:="First", WhatRow:=0, MinWidth:=0
FlatBar1.AddBandsToRebar FlatBar2.GetToolbarHwnd, "", 0, 0
'Don't need Design Time Container since we
'are adding them to the rebar so hide in case
'the form color changes
FlatBar1.Visible = False
FlatBar2.Visible = False
 

End Sub


Private Sub Form_Resize()
FlatBar1.ResizeRebar Me
End Sub


