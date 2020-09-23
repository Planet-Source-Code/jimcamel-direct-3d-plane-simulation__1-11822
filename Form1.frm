VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3192
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pi As Integer
Dim DXMain As New DirectX7
Dim DDMain As DirectDraw4
Dim RMMain As Direct3DRM3
Dim DIMain As DirectInput
Dim DIDevice As DirectInputDevice
Dim DIState As DIKEYBOARDSTATE
Dim DSPrim As DirectDrawSurface4
Dim DSBack As DirectDrawSurface4
Dim SDPrim As DDSURFACEDESC2
Dim DDBack As DDSCAPS2
Dim RMDevice As Direct3DRMDevice3
Dim RMView As Direct3DRMViewport2
' Direct3D - Frames & Meshes - Frames are "containers" for the 3D meshes
Dim FRRoot As Direct3DRMFrame3
Dim FRLight As Direct3DRMFrame3
Dim FRCam As Direct3DRMFrame3
Dim FRShip As Direct3DRMFrame3
Dim MSShip As Direct3DRMMeshBuilder3
Dim FRGround As Direct3DRMFrame3      ' Frame containing the ship
Dim MSGround As Direct3DRMMeshBuilder3 ' Mesh containing the ship
Dim LTMain As Direct3DRMLight
Dim LTSpot As Direct3DRMLight
Dim dXAng As Single
Dim dYAng As Single
Dim dZAng As Single
Dim iXPos As Integer
Dim iYPos As Integer
Dim iZPos As Integer

Sub DX_Init()
  Set DDMain = DXMain.DirectDraw4Create("")
  DDMain.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
  DDMain.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
  SDPrim.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  SDPrim.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
  SDPrim.lBackBufferCount = 1
  Set DSPrim = DDMain.CreateSurface(SDPrim)
  DDBack.lCaps = DDSCAPS_BACKBUFFER
  Set DSBack = DSPrim.GetAttachedSurface(DDBack)
  DSBack.SetForeColor RGB(255, 255, 255)
  Set RMMain = DXMain.Direct3DRMCreate()
  Set RMDevice = RMMain.CreateDeviceFromSurface("IID_IDirect3DRGBDevice", DDMain, DSBack, D3DRMDEVICE_DEFAULT)
  RMDevice.SetBufferCount 2
  RMDevice.SetQuality D3DRMRENDER_GOURAUD
  RMDevice.SetTextureQuality D3DRMTEXTURE_NEAREST
  RMDevice.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY
  Set DIMain = DXMain.DirectInputCreate()
  Set DIDevice = DIMain.CreateDevice("GUID_SysKeyboard")
  DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
  DIDevice.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  DIDevice.Acquire
End Sub

Sub DX_Scene()
  Set FRRoot = RMMain.CreateFrame(Nothing)
  Set FRLight = RMMain.CreateFrame(FRRoot)
  Set FRShip = RMMain.CreateFrame(FRRoot)
  Set FRGround = RMMain.CreateFrame(FRRoot)
  Set FRCam = RMMain.CreateFrame(FRRoot)
  FRRoot.SetSceneBackgroundRGB 0.4, 0.6, 1
  FRCam.SetPosition Nothing, 25, 50, 0
  Set RMView = RMMain.CreateViewport(RMDevice, FRCam, 0, 0, 640, 480)
  RMView.SetBack 300
  FRLight.SetPosition Nothing, 0, 25, 0
  Set LTSpot = RMMain.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
  FRLight.AddLight LTSpot
  Set LTMain = RMMain.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)
  FRRoot.AddLight LTMain
  
  Set MSShip = RMMain.CreateMeshBuilder()
  MSShip.LoadFromFile planename, 0, 0, Nothing, Nothing
  MSShip.ScaleMesh 0.1, 0.1, 0.1
  FRShip.AddVisual MSShip
  
  Set MSGround = RMMain.CreateMeshBuilder()
  MSGround.LoadFromFile App.Path & "\ground.x", 0, 0, Nothing, Nothing
  MSGround.ScaleMesh 0.01, 0.01, 0.01
  FRGround.AddVisual MSGround

  dXAng = 0
  dYAng = 0
  dZAng = 0
  iXPos = 0
  iYPos = 1
  iZPos = 2.5
End Sub

Sub DX_Render()
  On Local Error Resume Next
  Do
    DoEvents
    DX_Input
    RMView.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER
    RMDevice.Update
    DSBack.DrawText 200, 0, "l337 Planes By Adrian Clark", False
    DSBack.DrawText 140, 460, "Press [Esc] to exit or use the arrow keys to control the ship", False
    RMView.Render FRRoot
    DSPrim.Flip Nothing, DDFLIP_WAIT
  Loop

End Sub

Sub DX_Input()
  DIDevice.GetDeviceStateKeyboard DIState
  If DIState.Key(DIK_ESCAPE) <> 0 Then Call DX_Exit
  If DIState.Key(DIK_COMMA) <> 0 Then
    'Let dXAng = dXAng + 0.2
    Let dYAng = dYAng + 0.2
    'If dXAng > 1 Then Let dXAng = 1
  End If
  If DIState.Key(DIK_PERIOD) <> 0 Then
    'Let dXAng = dXAng - 0.2
    Let dYAng = dYAng - 0.2
    'If dXAng < -1 Then Let dXAng = -1
  End If
If DIState.Key(DIK_LEFT) <> 0 Then
    'Let dXAng = dXAng - 0.2
    Let dZAng = dZAng - 0.2
    'If dXAng < -1 Then Let dXAng = -1
  End If
  If DIState.Key(DIK_RIGHT) <> 0 Then
    'Let dXAng = dXAng - 0.2
    Let dZAng = dZAng + 0.2
    'If dXAng < -1 Then Let dXAng = -1
  End If
  
    
  'If DIState.Key(DIK_RIGHT) = 0 And DIState.Key(DIK_LEFT) = 0 Then
   ' If dXAng < 0 Then Let dXAng = dXAng + 0.2
   ' If dXAng > 0 Then Let dXAng = dXAng - 0.2
  'End If
 
  If DIState.Key(DIK_DOWN) <> 0 Then
    Let dXAng = dXAng - 0.2
    'Let iXPos = iXPos + (Sin(dYAng) * 2)
    'Let iYPos = iYPos + (Cos(dYAng) * 2)
  End If
  If DIState.Key(DIK_UP) <> 0 Then
    Let dXAng = dXAng + 0.2
      End If
  If DIState.Key(DIK_A) <> 0 Then
    Let iXPos = iXPos - (Cos((Pi / 2) - dZAng) * (Cos(dXAng) * 2))
    Let iZPos = iZPos - (Cos((Pi / 2) - dXAng) * 2)
    Let iYPos = iYPos - (Cos(dZAng) * (Cos(dXAng) * 2))
  End If
 If DIState.Key(DIK_Z) <> 0 Then
    Let iXPos = iXPos + (Cos((Pi / 2) - dZAng) * (Cos(dXAng) * 2))
    Let iZPos = iZPos + (Cos((Pi / 2) - dXAng) * 2)
    Let iYPos = iYPos + (Cos(dZAng) * (Cos(dXAng) * 2))
End If
    
    
  FRShip.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
  FRShip.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, dXAng + (Pi / 2)
  FRShip.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, dYAng
  FRShip.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, dZAng
  FRShip.SetPosition Nothing, iZPos + 10, iYPos, iXPos
  'FRCam.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, Pi
  FRCam.LookAt FRShip, Nothing, D3DRMCONSTRAIN_X
  FRGround.SetPosition Nothing, 0, 0, 0
  FRGround.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 0
  FRGround.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 0
  FRGround.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -1 * (Pi / 2)
  DSBack.DrawText 100, 460, dXAng & ", " & dYAng & ", " & dZAng, False
End Sub

Sub DX_Exit()
  Call DDMain.RestoreDisplayMode
  Call DDMain.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
  Call DIDevice.Unacquire
  End
End Sub

Private Sub Form_Load()
  Me.Show

  Pi = 3.14159265358979
  DoEvents
  DX_Init
  DX_Scene
  DX_Render
End Sub
