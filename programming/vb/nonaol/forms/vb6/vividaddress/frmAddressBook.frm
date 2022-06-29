VERSION 5.00
Begin VB.Form frmAddressBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adress Book"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "frmAddressBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdReadIni 
      Caption         =   "&Load Info"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdWriteToIni 
      Caption         =   "&Save Info"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "E-Mail"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Last Name"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "First Name"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strFirstName As String
Private strLastName As String
Private strEmail As String

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdReadIni_Click()
txtFirstName.Text = ReadFromIni("Address Book", "First Name", "c:\adressbook\adressbook.ini")
txtLastName.Text = ReadFromIni("Address Book", "Last Name", "c:\adressbook\adressbook.ini")
txtEmail.Text = ReadFromIni("Address Book", "Email", "c:\adressbook\adressbook.ini")



End Sub

Private Sub cmdWriteToIni_Click()
strFirstName = txtFirstName.Text
strLastName = txtLastName.Text
strEmail = txtEmail.Text
Call writetoini("Address Book", "First Name", strFirstName, "c:\adressbook\adressbook.ini")
Call writetoini("Address Book", "Last Name", strLastName, "c:\adressbook\adressbook.ini")
Call writetoini("Address Book", "Email", strEmail, "c:\adressbook\adressbook.ini")




End Sub
