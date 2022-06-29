Attribute VB_Name = "Module2"
' Fill the list of player positions.
Sub fillPoitionsList()
Dim position
    ' Ask the object application (SOCCER.MAK) for all the position names.
    For Each position In gServerApp.positions
        Form1.listPlayerPosition.AddItem position
    Next
End Sub

' Fill the list of teams in the list box with
' all the teams listed in SOCCER.MAK.
Sub fillTeamsList()
Dim team_name$
Dim loop1%

    ' List the total number of teams in the Number of Teams caption.
    Form1.labelNumberTeams.Caption = Trim(Str(gServerApp.Teams.Count - 1))

    ' Clear the list boxes.
    Form1.listTeams.Clear
    Form1.listPlayerTeam.Clear

    ' Note there are two lists to fill:
    ' 1. The list of teams for the graphical soccer field display.
    ' 2. The list of players in the individual player information section.
    For loop1% = 1 To gServerApp.Teams.Count
        ' Get the team name from the Teams collection.
        team_name$ = gServerApp.Teams.Item(loop1%).Name
        Form1.listTeams.AddItem team_name$
        Form1.listPlayerTeam.AddItem team_name$
    Next
   
End Sub

' Fill listPlayerName list box with
' all the teams listed in SOCCER.MAK.
Sub fillPlayersList()
Dim player_name$
Dim loop1%

    ' Get the number of players and display it in the Total # of Players caption.
    Form1.labelNumberPlayers.Caption = Trim(Str(gServerApp.Players.Count))

    'Clear the list boxes.
    Form1.listPlayerName.Clear
    
    For loop1% = 1 To gServerApp.Players.Count
        ' Get the team name from the collection.
        player_name$ = gServerApp.Players.Item(loop1%).Name
        Form1.listPlayerName.AddItem player_name$
    Next
   
End Sub

' Resets the names of all the positions on the form.
Sub ResetPositionNames()
    Form1.LabelStriker.Caption = "Striker"
    Form1.LabelCenter.Caption = "Center"
    Form1.labelRightForward.Caption = "Right Forward"
    Form1.LabelLeftForward.Caption = "Left Forward"
    Form1.LabelLeftWing.Caption = "Left Wing"
    Form1.LabelCenterMidfielder.Caption = "Center Midfielder"
    Form1.labelRightWing.Caption = "Right Wing"
    Form1.LabelLeftFullback.Caption = "Left Fullback"
    Form1.LabelSweeper.Caption = "Sweeper"
    Form1.LabelRightFullback.Caption = "Right Fullback"
    Form1.LabelGoalie.Caption = "Goalie"
End Sub

