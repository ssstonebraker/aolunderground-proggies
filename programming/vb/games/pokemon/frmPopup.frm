VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menua 
      Caption         =   "MenuA"
      Begin VB.Menu delete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu MenuB 
      Caption         =   "MenuB"
      Begin VB.Menu bottomcenter 
         Caption         =   "Bottom - Center"
      End
      Begin VB.Menu topc 
         Caption         =   "Top - Center"
      End
      Begin VB.Menu ontop 
         Caption         =   "Place Me On Top"
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bottomcenter_Click()
    frmChatroom.Top = (Screen.Height - frmChatroom.Height) - 100
    frmChatroom.Left = (Screen.Width - frmChatroom.Width) / 2
End Sub
Private Sub delete_Click()
    If frmSplash.lstPlayers.ListIndex = -1 Then
        MsgBoxA Me, "Select a game to delete first!"
    Else
        Kill PathA & "\" & LCase(TrimSpaces(frmSplash.lstPlayers.Text))
        SaveSetting "Pokémon Adventure", "Introduction", frmSplash.lstPlayers.List(frmSplash.lstPlayers.ListIndex), Empty
        SaveSetting "Pokémon Adventure", "Introduction", frmSplash.lstPlayers.Text, Empty
        frmSplash.lstPlayers.RemoveItem frmSplash.lstPlayers.ListIndex
        SaveGames frmSplash.lstPlayers
    End If
End Sub
Private Sub ontop_Click()
    FormOnTop frmChatroom
End Sub
Private Sub topc_Click()
    frmChatroom.Top = 100
    frmChatroom.Left = (Screen.Width - frmChatroom.Width) / 2
End Sub
