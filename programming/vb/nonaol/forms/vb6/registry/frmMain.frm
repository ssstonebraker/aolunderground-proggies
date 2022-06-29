VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plastik's System Registry Example"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   650
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   3615
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdReset_Click()
'delete the setting
Call DeleteSetting("Registry Example", "Times Loaded", "Setting")
End Sub

Private Sub Form_Load()
Dim intLoaded As Integer

 On Error Resume Next 'if there is an error because its empty, dont stop!
 
 'get the setting from the system registry
 intLoaded% = GetSetting("Registry Example", "Times Loaded", "Setting")
 'set the intLoaded% variable with the setting it gets from the
 'registry, Registry Example is the Application name that it will
 'go under, and Times Loaded is the section that is located in the
 'appname directory and Setting is the Key that holds the setting
 'in the directories
 
 intLoaded% = intLoaded% + 1 'add up 1 everytime the user loads it
 Call SaveSetting("Registry Example", "Times Loaded", "Setting", intLoaded%)
 'saves the setting to the registry after it has added 1 to it!
 
 'set the labels caption
 lblDisplay.Caption = "This program has been loaded " & intLoaded% & " time(s)!"
 
End Sub

