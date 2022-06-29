VERSION 2.00
Begin Form frmReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bugs Report"
   ClientHeight    =   4350
   ClientLeft      =   2355
   ClientTop       =   2655
   ClientWidth     =   4545
   Height          =   4755
   Left            =   2295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4545
   Top             =   2310
   Width           =   4665
   Begin CommonDialog dlgPrint 
      Left            =   3345
      Top             =   2310
   End
   Begin PictureBox Picture1 
      Height          =   4185
      Left            =   60
      ScaleHeight     =   4155
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   90
      Width           =   4425
      Begin CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   615
         Left            =   2775
         TabIndex        =   3
         Top             =   3405
         Width           =   1380
      End
      Begin CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   615
         Left            =   2775
         TabIndex        =   2
         Top             =   90
         Width           =   1380
      End
      Begin ListBox lstSystems 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   3630
         Left            =   75
         TabIndex        =   1
         Top             =   90
         Width           =   2550
      End
   End
End
Option Explicit

Sub cmdCancel_Click ()

    Unload frmReport

End Sub

Sub cmdPrint_Click ()

    Dim nCurrentCopy As Integer
    Dim sOldRecordsource As String

    Dim nC2 As Integer, nC3 As Integer, nC4 As Integer

    Const PD_DISABLEPRINTTOFILE = &H80000
    Const PD_HIDEPRINTTOFILE = &H100000
    Const PD_NOPAGENUMS = &H8&
    Const PD_NOSELECTION = &H4&


    dlgPrint.Flags = PD_DISABLEPRINTTOFILE + PD_HIDEPRINTTOFILE + PD_NOPAGENUMS + PD_NOSELECTION
    dlgPrint.Action = 5

    screen.MousePointer = 11

    sOldRecordsource = frmMainForm!datBugs.RecordSource

    frmMainForm!datBugs.RecordSource = "Select * from Bugs where Bugs.System_ID = " & lstSystems.ItemData(lstSystems.ListIndex) & " and not cleared"
    frmMainForm!datBugs.Refresh
    

    For nCurrentCopy = 1 To dlgPrint.Copies

        printer.FontSize = 18
        printer.Print "Summary of uncleared bugs as of " & Now
        printer.Line -Step(printer.TextWidth("Summary of uncleared bugs as of " & Now), 0)

        printer.FontSize = 13

        printer.Print
        printer.Print

        
        frmMainForm!datSystems.Recordset.FindFirst "System_ID" = lstSystems.ItemData(lstSystems.ListIndex)
        printer.FontItalic = True
        printer.Print "System    :";
        printer.FontItalic = False
        printer.Print frmMainForm!datSystems.Recordset.Fields("System_Name") & "  (" & frmMainForm!datSystems.Recordset.Fields("Makefile") & ")"
        printer.Print
        printer.Print
        
        printer.FontBold = True
        printer.FontUnderline = True
        printer.Print "ID    Description                               Created         "
        
        nC2 = printer.TextWidth("ID    ")
        nC3 = printer.TextWidth("ID    Description                               ")

        printer.FontUnderline = False
        printer.FontBold = False
        printer.Print

        On Error Resume Next
        frmMainForm.datBugs.Recordset.MoveFirst
        On Error GoTo 0
        Do Until frmMainForm.datBugs.Recordset.EOF
        
            printer.CurrentX = 0
            printer.Print frmMainForm.datBugs.Recordset.Fields("Bug_ID");

            printer.CurrentX = nC2
            printer.Print frmMainForm.datBugs.Recordset.Fields("Description");

            printer.CurrentX = nC3
            printer.Print frmMainForm.datBugs.Recordset.Fields("CreatedOn")

            frmMainForm.datBugs.Recordset.MoveNext
        
        Loop

        printer.EndDoc

    Next

    frmMainForm!datBugs.RecordSource = sOldRecordsource
    frmMainForm!datBugs.Refresh

    screen.MousePointer = 0

End Sub

Sub Form_Load ()

    Dim nIndex As Integer

    frmReport.Left = (screen.Width - frmReport.Width) / 2
    frmReport.Top = (screen.Height - frmReport.Height) / 2
    
    
    lstSystems.Clear

    For nIndex = 1 To frmMainForm!cboSystemName.ListCount - 1

        lstSystems.AddItem frmMainForm!cboSystemName.List(nIndex)
        lstSystems.ItemData(lstSystems.ListCount - 1) = frmMainForm!cboSystemName.ItemData(nIndex)

    Next

    lstSystems.ListIndex = 0

End Sub

