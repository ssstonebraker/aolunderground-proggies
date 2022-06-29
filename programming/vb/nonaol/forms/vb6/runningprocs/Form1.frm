VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processes Running on Windows9x and Windows 2000"
   ClientHeight    =   3495
   ClientLeft      =   2835
   ClientTop       =   2625
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdProcs 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FillStyle       =   1
      AllowUserResizing=   1
      FormatString    =   ""
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_PATH = 260
Const TH32CS_SNAPPROCESS = 2&

Private Type PROCESSENTRY32
    lSize            As Long
    lUsage           As Long
    lProcessId       As Long
    lDefaultHeapId   As Long
    lModuleId        As Long
    lThreads         As Long
    lParentProcessId As Long
    lPriClassBase    As Long
    lFlags           As Long
    sExeFile         As String * MAX_PATH
End Type

Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" _
    Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, _
    ByVal lProcessId As Long) As Long
    
Private Declare Function ProcessFirst Lib "kernel32" _
    Alias "Process32First" (ByVal hSnapshot As Long, _
    uProcess As PROCESSENTRY32) As Long
    
Private Declare Function ProcessNext Lib "kernel32" _
    Alias "Process32Next" (ByVal hSnapshot As Long, _
    uProcess As PROCESSENTRY32) As Long
Private Sub Form_Load()
Dim sExeName   As String
Dim sPid       As String
Dim sParentPid As String
Dim lSnapShot  As Long
Dim r          As Long
Dim uProcess   As PROCESSENTRY32

lSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If lSnapShot <> 0 Then
    With grdProcs
    .Clear
    .Rows = 1
    .TextMatrix(0, 0) = "Module Name"
    .TextMatrix(0, 1) = "Process Id"
    .TextMatrix(0, 2) = "Parent" & vbCrLf & "Process"
    .TextMatrix(0, 3) = "Threads"
    .RowHeight(0) = 400
    .ColWidth(0) = 4200
    .ColWidth(1) = 950
    .ColWidth(2) = 950
    .ColWidth(3) = 775
    .ColAlignment(0) = flexAlignLeftBottom
    .ColAlignment(1) = flexAlignLeftBottom
    .ColAlignment(2) = flexAlignLeftBottom
    .ColAlignment(3) = flexAlignLeftBottom
    
    uProcess.lSize = Len(uProcess)
    r = ProcessFirst(lSnapShot, uProcess)

    Do While r
        sExeName = Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1)
        sPid = Hex$(uProcess.lProcessId)
        sParentPid = Hex$(uProcess.lParentProcessId)
        .AddItem sExeName & vbTab & sPid & vbTab & _
                sParentPid & vbTab & CStr(uProcess.lThreads)
        r = ProcessNext(lSnapShot, uProcess)
    Loop
    CloseHandle (lSnapShot)
    End With
End If
End Sub


