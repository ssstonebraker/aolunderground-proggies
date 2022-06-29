VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cpu boot"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2055
   Icon            =   "cpu_fire_up.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleMode       =   0  'User
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame f 
      Caption         =   "cpu was fired up:"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Label Label4 
         Caption         =   "minutes ago"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "hours and"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'cpu boot example by sic

'this is a very very simple example on how
'to get the length of time since your last
're-boot or cold boot.

'this file should have been downloaded from:
'http://www.knk2000.com/knk     OR:
'directly from AOL's boards

'if not email me with:
'      the url you downloaded it from
'      your handle/nickanme

'and i'll get it straighten out

'contacts:
    
    'email: codis1@hotmail.com, isickening@aol.com
    '  aim: ms visual cpp, o8g, v8i
    '  icq: i hate it...dont use it =X


Private Sub Form_Load()
stayontop Me 'get on top
Call lastboot(Label1, Label3) ' call the sub
End Sub

