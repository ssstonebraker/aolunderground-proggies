VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   1950
   ClientTop       =   2805
   ClientWidth     =   7365
   Height          =   3840
   Left            =   1890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7365
   Top             =   2460
   Width           =   7485
   Begin TextBox Text5 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2265
      TabIndex        =   9
      Top             =   1875
      Width           =   4440
   End
   Begin TextBox Text4 
      DataField       =   "Au_ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2265
      TabIndex        =   8
      Top             =   1290
      Width           =   4440
   End
   Begin TextBox Text3 
      DataField       =   "Year Published"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2265
      TabIndex        =   7
      Top             =   705
      Width           =   4440
   End
   Begin TextBox Text2 
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2265
      TabIndex        =   6
      Top             =   2460
      Width           =   4440
   End
   Begin TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2265
      TabIndex        =   5
      Top             =   120
      Width           =   4440
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "BIBLIO.MDB"
      Exclusive       =   0   'False
      Height          =   270
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Titles"
      Top             =   2910
      Width           =   6720
   End
   Begin Label Label5 
      Caption         =   "Author"
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   1335
      Width           =   1515
   End
   Begin Label Label4 
      Caption         =   "Publisher"
      Height          =   210
      Left            =   600
      TabIndex        =   3
      Top             =   2475
      Width           =   1515
   End
   Begin Label Label3 
      Caption         =   "ISBN"
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   1905
      Width           =   1515
   End
   Begin Label Label2 
      Caption         =   "Date Published"
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   765
      Width           =   1515
   End
   Begin Label Label1 
      Caption         =   "Title"
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   195
      Width           =   1515
   End
End
Option Explicit

Sub Form_Load ()

    ' We need to load up the databasename into the data control. This
    ' code assumes that the database is in the same directory as the
    ' application.
    data1.DatabaseName = app.Path & "\" & "biblio.mdb"

End Sub

