VERSION 2.00
Begin Form rfmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Using the Move methods"
   ClientHeight    =   4245
   ClientLeft      =   1230
   ClientTop       =   2505
   ClientWidth     =   7365
   Height          =   4650
   Left            =   1170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7365
   Top             =   2160
   Width           =   7485
   Begin CommandButton cmdLast 
      Caption         =   "Last"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   5895
      TabIndex        =   14
      Top             =   3630
      Width           =   1350
   End
   Begin CommandButton cmdNext 
      Caption         =   "Next"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   3975
      TabIndex        =   13
      Top             =   3630
      Width           =   1350
   End
   Begin CommandButton cmdPrev 
      Caption         =   "Prev"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   2055
      TabIndex        =   12
      Top             =   3630
      Width           =   1350
   End
   Begin CommandButton cmdFirst 
      Caption         =   "First"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   135
      TabIndex        =   11
      Top             =   3630
      Width           =   1350
   End
   Begin PictureBox Picture1 
      Height          =   3450
      Left            =   135
      ScaleHeight     =   3420
      ScaleWidth      =   7065
      TabIndex        =   5
      Top             =   75
      Width           =   7095
      Begin Data datTitles 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "biblio.mdb"
         Exclusive       =   0   'False
         Height          =   285
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Titles"
         Top             =   2835
         Visible         =   0   'False
         Width           =   6555
      End
      Begin TextBox Text1 
         BackColor       =   &H00FFFF00&
         DataField       =   "Title"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2145
         TabIndex        =   10
         Top             =   315
         Width           =   4140
      End
      Begin TextBox Text2 
         BackColor       =   &H00FFFF00&
         DataField       =   "Year Published"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2145
         TabIndex        =   9
         Top             =   810
         Width           =   4140
      End
      Begin TextBox Text3 
         BackColor       =   &H00FFFF00&
         DataField       =   "Au_ID"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2145
         TabIndex        =   8
         Top             =   1320
         Width           =   4140
      End
      Begin TextBox Text4 
         BackColor       =   &H00FFFF00&
         DataField       =   "ISBN"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2145
         TabIndex        =   7
         Top             =   1830
         Width           =   4140
      End
      Begin TextBox Text5 
         BackColor       =   &H00FFFF00&
         DataField       =   "PubID"
         DataSource      =   "datTitles"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2145
         TabIndex        =   6
         Top             =   2340
         Width           =   4140
      End
      Begin Label Label1 
         Caption         =   "Title"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   615
         TabIndex        =   0
         Top             =   300
         Width           =   1425
      End
      Begin Label Label2 
         Caption         =   "Date Published"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   450
         TabIndex        =   1
         Top             =   840
         Width           =   1680
      End
      Begin Label Label3 
         Caption         =   "Author"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   615
         TabIndex        =   2
         Top             =   1350
         Width           =   1425
      End
      Begin Label Label4 
         Caption         =   "ISBN"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   615
         TabIndex        =   3
         Top             =   1875
         Width           =   1425
      End
      Begin Label Label5 
         Caption         =   "Publisher"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   615
         TabIndex        =   4
         Top             =   2400
         Width           =   1425
      End
   End
End
Option Explicit

Sub cmdFirst_Click ()

    datTitles.Recordset.MoveFirst

End Sub

Sub cmdLast_Click ()

    datTitles.Recordset.MoveLast

End Sub

Sub cmdNext_Click ()

    If Not datTitles.Recordset.EOF Then datTitles.Recordset.MoveNext

End Sub

Sub cmdPrev_Click ()

    If Not datTitles.Recordset.BOF Then datTitles.Recordset.MovePrevious

End Sub

Sub Data1_Validate (Action As Integer, Save As Integer)

    Const DATA_ACTIONCANCEL = 0
    Const DATA_ACTIONUPDATE = 6
    
    If Action = DATA_ACTIONUPDATE Then Action = DATA_ACTIONCANCEL

End Sub

Sub datTitles_Validate (Action As Integer, Save As Integer)

    Save = False

End Sub

