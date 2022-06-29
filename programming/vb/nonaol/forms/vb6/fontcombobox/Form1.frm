VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fonts in a combo box"
   ClientHeight    =   2145
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6690
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1920
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "(select font)"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    Label1.Font = Combo1.Text
    Label1.Caption = Combo1.Text
End Sub

Private Sub Form_Load()
    For X = 0 To Printer.FontCount
        Combo1.AddItem Printer.Fonts(X)
    Next X
    Combo1.RemoveItem (0)
    Combo1.ListIndex = 0
End Sub
