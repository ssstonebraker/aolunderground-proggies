VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
'Original Code By Hank Marquis
'Modifications By BitStream(JKWild)
'June 1997
'----------------------------------------------------
'API to create, Active(resume), suspend, and end Threads
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'API to set/get priority of threads
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

'member of class which is the handle to the thread
Public Thread As Long
'member of class which indicates if the thread is
'   suspended (False) or resumed (True)
Public ThreadStatus As Boolean
'use this option to create a thread which is "suspended" at first
Private Const CREATE_SUSPENDED = &H4
'Priority Constants
Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT


Public Sub Initialize(lpfnBasFunc As Long)
    Dim dwStackSize As Long
    Dim dwCreationFlags As Long
    Dim lpThreadId As Long
    Dim lpParameter As Long
    Dim myNull As Long

    myNull = 0&                         'create a "Null" pointer
    dwStackSize = 0                     '0 means use exe stack size
    dwCreationFlags = CREATE_SUSPENDED  'initially suspended
    Me.Thread = CreateThread(myNull, dwStackSize, lpfnBasFunc, myNull, dwCreationFlags, lpThreadId)
    If Me.Thread = myNull Then MsgBox "Create thread failed!"
End Sub

Public Property Get Enabled() As Boolean
    Enabled = Me.ThreadStatus
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    'make sure we don't try to start|stop a started|stopped thread
    If vNewValue = True And Me.ThreadStatus = False Then
        ResumeThread Me.Thread 'resume
        Me.ThreadStatus = True
    ElseIf Me.ThreadStatus = True Then
        SuspendThread Me.Thread    'suspend
        Me.ThreadStatus = False
    End If
End Property


Public Property Get Priority() As Long
    Priority = GetThreadPriority(Me.Thread)
End Property

Public Property Let Priority(ByVal vNewValue As Long)
    Select Case vNewValue
        Case -2
            SetThreadPriority Me.Thread, THREAD_PRIORITY_LOWEST
        Case -1
            SetThreadPriority Me.Thread, THREAD_PRIORITY_BELOW_NORMAL
        Case 0
            SetThreadPriority Me.Thread, THREAD_PRIORITY_NORMAL
        Case 1
            SetThreadPriority Me.Thread, THREAD_PRIORITY_ABOVE_NORMAL
        Case 2
            SetThreadPriority Me.Thread, THREAD_PRIORITY_HIGHEST
    End Select
End Property

Public Function Term()
    Dim ExitCode As Long
    Dim result As Integer
    
    result = GetExitCodeThread(Me.Thread, ExitCode)
    ExitThread ExitCode
    result = result + CloseHandle(Me.Thread)
    'Non Zero Result indicates Success
    'Remember to Terminate Class following this Function
End Function
