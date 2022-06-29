VERSION 2.00
Begin Form frmMain 
   Caption         =   "Adding and removing items with grids"
   ClientHeight    =   5355
   ClientLeft      =   1620
   ClientTop       =   2535
   ClientWidth     =   7290
   Height          =   5760
   Left            =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7290
   Top             =   2190
   Width           =   7410
   Begin CommandButton cmdRemove 
      Caption         =   "&Remove Item"
      Height          =   450
      Left            =   5430
      TabIndex        =   2
      Top             =   4800
      Width           =   1605
   End
   Begin CommandButton cmdAdd 
      Caption         =   "&Add Item"
      Height          =   450
      Left            =   225
      TabIndex        =   1
      Top             =   4800
      Width           =   1605
   End
   Begin Grid grdContacts 
      FixedCols       =   0
      Height          =   4455
      Left            =   225
      TabIndex        =   0
      Top             =   255
      Width           =   6780
   End
End
Option Explicit

Sub cmdAdd_Click ()

    Dim sColumn1 As String
    Dim sColumn2 As String

    sColumn1 = InputBox("Enter the contact's forename")
    sColumn2 = InputBox("Enter the contact's surname")

    grdContacts.AddItem sColumn1 & Chr$(9) & sColumn2, grdContacts.Rows - 1

End Sub

Sub cmdRemove_Click ()

    grdContacts.RemoveItem grdContacts.Row
    
End Sub

Sub Form_Load ()

    ' First we set the left and top properties of the form, to centre it on screen
    frmMain.Left = (Screen.Width - frmMain.Width) / 2
    frmMain.Top = (Screen.Height - frmMain.Height) / 2


    ' Next step is to set up the column widths of the columns in the grid
    grdContacts.ColWidth(0) = grdContacts.Width / 2
    grdContacts.ColWidth(1) = (grdContacts.Width / 2) - 10

    ' Finally, we can set up the column titles
    grdContacts.Row = 0
    grdContacts.Col = 0
    grdContacts.Text = "Contact's Forenames"
    
    grdContacts.Col = 1
    grdContacts.Text = "Contact's Surname"

End Sub
                   

