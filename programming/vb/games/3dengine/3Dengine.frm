VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3D engine"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "3Dengine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim g_dx As New DirectX7
Dim m_dd As DirectDraw7
Dim m_ddClipper As DirectDrawClipper
Dim m_rm As Direct3DRM3

Dim m_rmDevice As Direct3DRMDevice3
Dim m_rmViewport As Direct3DRMViewport2

Dim m_rootFrame As Direct3DRMFrame3
Dim m_lightFrame As Direct3DRMFrame3
Dim m_cameraFrame As Direct3DRMFrame3
Dim m_objectFrame As Direct3DRMFrame3
Dim BlocksFrame(10, 10) As Direct3DRMFrame3
Dim m_uvFrame As Direct3DRMFrame3
Dim m_light As Direct3DRMLight
Dim m_meshBuilder As Direct3DRMMeshBuilder3
Dim m_object As Direct3DRMMeshBuilder3

Dim m_width As Long
Dim m_height As Long
Dim m_running As Boolean
Dim m_finished As Boolean

Dim Map_Data(10, 10) As Integer

Private Sub Form_Load()
        Show
        DoEvents
        InitRM
        FindMediaDir "egg.x"
        InitScene
        RenderLoop
        CleanUp
        End
End Sub

Sub CleanUp()
    m_running = False
    
    Exit Sub
    Set m_light = Nothing
    Set m_meshBuilder = Nothing
    Set m_object = Nothing

    Set m_lightFrame = Nothing
    Set m_cameraFrame = Nothing
    Set m_objectFrame = Nothing
    
    Dim X As Integer, Y As Integer
    For X = 0 To 10: For Y = 0 To 10
        Set BlocksFrame(X, Y) = Nothing
    Next Y: Next X
    Set m_rootFrame = Nothing

    Set m_rmDevice = Nothing
    Set m_ddClipper = Nothing
    Set m_rm = Nothing
    Set m_dd = Nothing
 
End Sub

Sub InitRM()


    'Create Direct Draw From Current Display Mode
    Set m_dd = g_dx.DirectDrawCreate("")
    
    'Create new clipper object and associate it with a window'
    Set m_ddClipper = m_dd.CreateClipper(0)
    m_ddClipper.SetHWnd Picture1.hWnd
        
    
    'save the widht and height of the picture in pixels
    m_width = Picture1.ScaleWidth
    m_height = Picture1.ScaleHeight
    
    'Create the Retained Mode object
    Set m_rm = g_dx.Direct3DRMCreate()

    
    'Create the Retained Mode device to draw to
    Set m_rmDevice = m_rm.CreateDeviceFromClipper(m_ddClipper, "", m_width, m_height)
    
    m_rmDevice.SetQuality D3DRMRENDER_GOURAUD
    
End Sub

Sub InitScene()
Dim tempNum As String, X As Integer, Y As Integer
Open App.Path & "\Data.txt" For Input As #1
    For Y = 0 To 10
    For X = 0 To 10
        Line Input #1, tempNum
        Map_Data(X, Y) = Val(tempNum)
    Next X
    Next Y
Close #1

    'Setup a scene graph with a camera light and object
    Set m_rootFrame = m_rm.CreateFrame(Nothing)
    Set m_cameraFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_lightFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_objectFrame = m_rm.CreateFrame(m_rootFrame)
    For X = 0 To 10: For Y = 0 To 10
        Set BlocksFrame(X, Y) = m_rm.CreateFrame(m_rootFrame)
    Next Y: Next X
    'position the camera and create the Viewport
    'provide the device thre viewport uses to render, the frame whose orientation and position
    'is used to determine the camera, and a rectangle describing the extents of the viewport
    m_cameraFrame.SetPosition Nothing, 0, 15, -10
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_cameraFrame, 0, 0, m_width, m_height)
    
    
    'create a white light and hang it off the light frame
    Set m_light = m_rm.CreateLight(D3DRMLIGHT_POINT, RGB(199, 199, 199))
    m_lightFrame.AddLight m_light
    Set m_light = m_rm.CreateLight(D3DRMLIGHT_AMBIENT, RGB(30, 30, 30))
    m_lightFrame.AddLight m_light
    
    'For this sample we will load x files with geometry only
    'so create a meshbuilder object
    Set m_meshBuilder = m_rm.CreateMeshBuilder()
    m_meshBuilder.LoadFromFile "egg.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing

    'Set the scale
    m_meshBuilder.ScaleMesh 1, 1, 1
    
    'add the meshbuilder to the scene graph
    m_objectFrame.AddVisual m_meshBuilder
    
    For Y = 0 To 10
    For X = 0 To 10
        If Map_Data(X, Y) = 1 Then
            Call Load_Mesh("Box.x", 1, 1, 1, X, Y)
        Else
            Call Load_Mesh("Box.x", 1, 0.4, 1, X, Y)
        End If
        
        BlocksFrame(X, Y).SetPosition m_rootFrame, X * 2, 0, Y * 2
    Next X
    Next Y
    
    'Have the object rotating
    m_objectFrame.SetRotation Nothing, 0, 1, 0, 0.01
 '   m_lightFrame.SetRotation Nothing, 0, 1, 0, 0.01
    m_lightFrame.SetPosition m_rootFrame, 8, 10, 8
    m_objectFrame.SetPosition m_rootFrame, 6, 1, 10
    m_cameraFrame.LookAt BlocksFrame(5, 5), m_rootFrame, D3DRMCONSTRAIN_Z
    
End Sub

Sub Load_Mesh(sMesh As String, scale_X As Single, scale_Y As Single, scale_Z As Single, X As Integer, Y As Integer)
    'For this sample we will load x files with geometry only
    'so create a meshbuilder object
    Set m_meshBuilder = m_rm.CreateMeshBuilder()
    m_meshBuilder.LoadFromFile sMesh, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing

    'Set the scale
    m_meshBuilder.ScaleMesh scale_X, scale_Y, scale_Z
    
    'add the meshbuilder to the scene graph
    BlocksFrame(X, Y).AddVisual m_meshBuilder

End Sub

Sub RenderLoop()
    Dim t1 As Long
    Dim t2 As Long
    
    Dim delta As Single
    On Local Error Resume Next
    m_running = True
    t1 = g_dx.TickCount()
    Do While m_running = True
        t2 = g_dx.TickCount()
        delta = (t2 - t1) / 10
        t1 = t2
        m_rootFrame.Move delta  'increment velocities
        m_rmViewport.Clear D3DRMCLEAR_ALL    'clear the rendering surface rectangle described by the viewport
        m_rmViewport.Render m_rootFrame 'render to the device
        FixFloat
        m_rmDevice.Update   'blt the image to the screen
        DoEvents    'allows events to be processed even though we are in a tight loop
    Loop
End Sub

Sub FixFloat()
    On Local Error Resume Next
    Dim l As Single
    l = 6
    
    l = l / 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CleanUp
    End
End Sub


Private Sub Picture1_Paint()
    On Local Error Resume Next
    m_rmDevice.HandlePaint Picture1.hDC
End Sub

Sub FindMediaDir(sFile As String)
    On Local Error Resume Next
    If Mid$(App.Path, 2, 1) = ":" Then
        ChDrive Mid$(App.Path, 1, 1)
    End If
    ChDir App.Path
    If Dir$(sFile) = "" Then
        ChDir App.Path & "\3dx"
    End If
'    If Dir$(sFile) = "" Then
'        ChDir "..\..\3dx"
'    End If
End Sub

