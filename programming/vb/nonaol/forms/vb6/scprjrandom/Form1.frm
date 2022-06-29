VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "prjRandom"
   ClientHeight    =   2160
   ClientLeft      =   7875
   ClientTop       =   6795
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   2655
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Form1.frx":0000
      Left            =   1440
      List            =   "Form1.frx":0022
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###################################################
'#              Randomizing Example                #
'#                       By:                       #
'#                     Shahama                     #
'# Alright. In this example I put some colors in a #
'#    ListBox. Then I made a variable that was     #
'#  Randomized. It then selected the random color  #
'#                in the ListBox.                  #
'###################################################

Private Sub Command1_Click()
        x = Int(Rnd * List1.ListCount) ' Get a Random number between 1 and the number of things in the ListBox.
        List1.ListIndex = x  ' Select the color that matches the number.
End Sub

Private Sub Command2_Click()
    List1.AddItem (InputBox("Yodel. Put a thingy here.", "Yodel")) ' Ask person to add Item and then Add it.
End Sub

Private Sub Command3_Click()
     List1.RemoveItem List1.ListIndex 'Find selected Item and remove it.
End Sub

Private Sub Command4_Click()
    frmAbout.Show ' Show my about form
End Sub

Private Sub Command5_Click()
    End ' End my program
End Sub
