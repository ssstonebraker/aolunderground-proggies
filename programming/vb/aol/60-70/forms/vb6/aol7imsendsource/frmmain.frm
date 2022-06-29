VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aol 7.0 Send I.M by source"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Send Instant Message Example 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "s e n d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "this is the pre-set example, this is a set screen name and fixed message with the use of quotes."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Instant Message Example 1 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "s e n d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txt2_msg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txt1_sn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "this will take the text in the textbox's and use the module coding to insert them into the aol textbox's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Message/Body Of Text:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Send To (Screen Name) : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "This example and module were created by : source (aim: ciasource) and check out planetsourcecode for more examples"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3120
      Width           =   4095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Aol Version 7.0 Send Instant Message Example By Source
'This is fully commented and aimed to help you better understand
'Public Declare Functions and Public Const, something that isnt usually understood
'just copied.
'website that might be up but might not : www.dazedworld.com
'contact aim: ciasource

Private Sub Command1_Click()
'Using the vb command "call" you can extract coding from the module.
'In this case it uses the Public Function SendInstantMessage. Then
'where the txtsn as string and txtmsg as string are filled inside the
'module/function its filled it with txt1_sn.text, and txt2_msg.text
Call SendInstantMessage(txt1_sn.text, txt2_msg.text)
End Sub

Private Sub Command2_Click()
'Using the vb command "call" you can extract coding from the module.
'In this case it uses the Public Function SendInstantMessage.
'Then with the quotes it fills in the fixed screen name (ciasource) and the
'fixed message (hey source, im using your example)
Call SendInstantMessage("ciasource", "hey source, im using your example")
End Sub
