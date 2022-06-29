VERSION 2.00
Begin Form frmMainForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Common Controls Example"
   ClientHeight    =   5160
   ClientLeft      =   405
   ClientTop       =   1995
   ClientWidth     =   9135
   FillColor       =   &H00FFFFFF&
   Height          =   5565
   Left            =   345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   9135
   Top             =   1650
   Width           =   9255
   Begin Frame fraFilenameAndOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filename and options"
      ForeColor       =   &H00FF0000&
      Height          =   3000
      Left            =   30
      TabIndex        =   12
      Top             =   135
      Width           =   4395
      Begin CheckBox chkAnimate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Animate image control"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   330
         TabIndex        =   5
         Top             =   1785
         Width           =   2925
      End
      Begin CommandButton cmdQuit 
         Caption         =   "&Quit"
         Height          =   390
         Left            =   3045
         TabIndex        =   9
         Top             =   2340
         Width           =   1035
      End
      Begin OptionButton optShadow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shadow &both"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   4
         Top             =   1365
         Width           =   1815
      End
      Begin OptionButton optShadow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shadow other frame"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Top             =   1080
         Width           =   2025
      End
      Begin OptionButton optShadow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shadow this frame"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   750
         Width           =   2115
      End
      Begin CommandButton cmdLoad 
         Caption         =   "&Load"
         Height          =   405
         Left            =   270
         TabIndex        =   1
         Top             =   2355
         Width           =   1035
      End
      Begin TextBox txtFilename 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1005
         TabIndex        =   0
         Top             =   405
         Width           =   3195
      End
      Begin Label lblFileName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filename"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   450
         Width           =   1230
      End
   End
   Begin Timer timTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4260
      Top             =   4020
   End
   Begin Frame fraFileControls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filename controls"
      ForeColor       =   &H00FF0000&
      Height          =   3000
      Left            =   4530
      TabIndex        =   11
      Top             =   135
      Width           =   4485
      Begin FileListBox filFilenames 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2565
         Left            =   75
         Pattern         =   "*.bmp;*.ico;*.wmf"
         TabIndex        =   6
         Top             =   255
         Width           =   1890
      End
      Begin DirListBox dirDirectories 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   2025
         TabIndex        =   7
         Top             =   270
         Width           =   2310
      End
      Begin DriveListBox drvDrives 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2055
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
      End
   End
   Begin Image imgPicture 
      Height          =   1440
      Left            =   720
      Tag             =   "+20"
      Top             =   3495
      Width           =   3525
   End
End
Option Explicit

Sub chkAnimate_Click ()

    ' As with the other controls, the click event occurs whenever the user points the mouse
    ' at the checkbox itself and clicks the left mouse button once. In the case of this checkbox
    ' when that happens we need to toggle (flip) the timer control on and off. This will either
    ' start or stop the graphic displayed from moving.

    timTimer.Enabled = Not timTimer.Enabled

End Sub

Sub cmdLoad_Click ()

    Const MB_OK = 0

    ' When the load button is clicked we need to do a number of things. First check that the
    ' filename in the text box actually specifies an existing file, and then tell the picture
    ' box to use that file.

    ' To check if a file exists, we can use the DIR function (the same as in DOS).
    ' If the DIR function = "" then that means that the file was not found

    If Dir(txtFileName.Text) = "" Then
        
        ' Display a message box saying that the file could not be found. The MB_OK parameter
        ' tells VB to display an "OK" message box - its a simple message box with a single
        ' button on it saying "OK"
        MsgBox "Sorry, that file does not exist", MB_OK, "Error"

    Else
        ' Otherwise...load the picture specified into the picture box. This can be done using
        ' the VB LoadPicture function.
        imgPicture.Picture = LoadPicture(txtFileName.Text)

    End If

End Sub

Sub cmdQuit_Click ()

    ' If the user clicks on the QUIT button then we really should unload the form. If you unload
    ' every form in a program then that has the same effect as stopping the program dead in
    ' its tracks. Since our program only has 1 form then we only need to unload that one.

    ' The form's QueryUnload event will ask the user whether or not they really want to quit
    ' before actually unloading itself.

    Unload Me

End Sub

Sub dirDirectories_Change ()

    ' This event occurs whenever the user selects a new directory from the list shown. The
    ' new path needs to be fed into the file list box so that both the directory and file list
    ' boxes work together.

    filFileNames.Path = dirDirectories.Path

End Sub

Sub drvDrives_Change ()

   Dim nResponse As Integer

    ' This event occurs whenever the user changes the drive shown in the drive list box.
    ' In here we need to feed the new drive to the directory list box, in order to get the
    ' two working together.

    dirDirectories.Path = drvDrives.Drive

    End Sub

Sub filFilenames_Click ()

    ' Unsurprisingly, this event occurs whenever the user clicks on a file name in the file
    ' list box. In the case of our application we want to fire the path and the filename into
    ' the text box displayed on the left hand side of the screen.

    txtFileName.Text = filFileNames.Path & "\" & filFileNames.FileName

End Sub

Sub filFilenames_DblClick ()

    ' If the user double clicks a file name in the list box then they are as good as saying
    ' "Load this picture". All we need to do here is run the Load command button's click code

    cmdLoad_Click

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    
    ' The QueryUnload event for a form is called just before the form closes down. In here
    ' we can place code which checks whether the form really should be closing down, and if
    ' it shouldn't, then we can stop it.
    


    ' Define a constant called MB_YESNO. This is used by the message box function to display
    ' a message box with two buttons in, one for Yes, the other for No.
    Const MB_YESNO = 4

    ' Define a constant to tell the message box to also display a STOP icon.
    Const MB_ICONSTOP = 16

    ' Define a constant called IDYES. This is the result of the msgbox function if the user
    ' clicks the yes button, ie If Msgbox(......) = IDYES then answer was yes.
    Const IDNO = 7



    If MsgBox("Do you really want to quit", MB_YESNO + MB_ICONSTOP, "Quit") = IDNO Then

        Cancel = True

    End If

End Sub

Sub optShadow_Click (index As Integer)

    ' When any of the option buttons are clicked, this event code runs. It checks which
    ' of the option buttons is currently set (ie Value = -1) and shades the frames on screen
    ' accordingly.


    ' If the value of the option button is 1 then add a shadow to the text box
    If optShadow(0).Value = True Or optShadow(2).Value = True Then
        Add_Drop_Shadow Me, fraFilenameAndOptions, &HC00000
    Else
        Add_Drop_Shadow Me, fraFilenameAndOptions, Val(frmMainForm.BackColor)
    End If

    If optShadow(1).Value = True Or optShadow(2).Value = True Then
        Add_Drop_Shadow Me, fraFileControls, &HC00000
    Else
        Add_Drop_Shadow Me, fraFileControls, Val(frmMainForm.BackColor)
    End If

End Sub

Sub timTimer_Timer ()

    imgPicture.Left = imgPicture.Left + Val(imgPicture.Tag)

    If imgPicture.Left < 200 Or imgPicture.Left > (frmMainForm.ScaleWidth - imgPicture.Width) Then
        imgPicture.Tag = 0 - Val(imgPicture.Tag)
    End If

End Sub

