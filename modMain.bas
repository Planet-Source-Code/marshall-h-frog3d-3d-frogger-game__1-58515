Attribute VB_Name = "modMain"
'----------------------------
'-  Module Code for Frog3D  -
'----------------------------

'DirectX
Public DirectX As New DirectX7
Public LookFrame As Direct3DRMFrame3
Public CameraFrame As Direct3DRMFrame3
Public CameraPos As D3DVECTOR
Public SceneFrame As Direct3DRMFrame3
Public DirectDraw As DirectDraw4
Public Clipper As DirectDrawClipper
Public D3DRM As Direct3DRM3
Public Device As Direct3DRMDevice3
Public Viewport As Direct3DRMViewport2

'Game Timer
Public NowTime As Long, Delay As Integer, GameDelay As Integer
Public LastTick As Long

'Lights
Public Light As Direct3DRMLight
Public LightFrame As Direct3DRMFrame3

'General Declarations
Public GameMode As Integer
Public Const GAME = 1
Public Const MENU = 2

Public Fullscreen As Boolean
Public ScreenX As Long
Public ScreenY As Long
Public UseFiltering As Boolean
Public UseFlat As Boolean
Public UseMusic As Boolean

'Camera Settings
Public CameraView As Integer
Public CamAngle As Double
Public Const CHASE = 0
Public Const INSIDE = 1
Public Const SKY = 2
Public Const FREE = 3
Public Const HIGH = 4

'Objects
Public TempFrame As Direct3DRMFrame3
Public SkyFrame As Direct3DRMFrame3
Public SkyMesh As Direct3DRMMeshBuilder3
Public SkyTexture As Direct3DRMTexture3

Public GroundFrame As Direct3DRMFrame3
Public GroundMesh As Direct3DRMMeshBuilder3
Public WaterFrame As Direct3DRMFrame3
Public WaterMesh As Direct3DRMMeshBuilder3

Public FrogFrame As Direct3DRMFrame3
Public FrogMesh(4) As Direct3DRMMeshBuilder3
Public SplatMesh As Direct3DRMMeshBuilder3

Public CarFrame(9) As Direct3DRMFrame3
Public CarMesh(9) As Direct3DRMMeshBuilder3
Public CarSpeed(9) As Double
Public CarPos(9) As D3DVECTOR

Public AppleFrame As Direct3DRMFrame3
Public AppleMesh As Direct3DRMMeshBuilder3

Public BeeMesh(1) As Direct3DRMMeshBuilder3

Public BeeFrame As Direct3DRMFrame3
Public AppleBeeFrame As Direct3DRMFrame3
Public BeePos As D3DVECTOR
Public BeeRaised As Boolean
Public AppleBeeRaised As Boolean

Public FlyMesh(1) As Direct3DRMMeshBuilder3
Public FlyPos As D3DVECTOR
Public FlyFrame As Direct3DRMFrame3
Public FlyRaised As Boolean

Public FrogCount As Integer
Public FrogDir As Integer
Public FrogPos As D3DVECTOR
Public Squishing As Boolean

'frog direction constants
Public Const DIRFRWD = 8
Public Const DIRBACK = 2
Public Const DIRLEFT = 4
Public Const DIRRGHT = 6

'Objects for the menu
Public PlayMesh As Direct3DRMMeshBuilder3
Public QuitMesh As Direct3DRMMeshBuilder3
Public TitleMesh As Direct3DRMMeshBuilder3
Public MenuFrogMesh As Direct3DRMMeshBuilder3
Public MenuFrogFrame As Direct3DRMFrame3
Public MFY As Double
Public PlayFrame As Direct3DRMFrame3
Public QuitFrame As Direct3DRMFrame3
Public TitleFrame As Direct3DRMFrame3
Public CreditFrame As Direct3DRMFrame3
Public CreditMesh As Direct3DRMMeshBuilder3

'score and energy bar
Public ScoreFrame As Direct3DRMFrame3
Public ScoreMesh As Direct3DRMMeshBuilder3
Public ScoreWheelFrame As Direct3DRMFrame3

Public WheelMesh As Direct3DRMMeshBuilder3
Public Score As Long

Public EnergyFrame As Direct3DRMFrame3
Public EnergyMesh As Direct3DRMMeshBuilder3
Public EnergyWheelFrame As Direct3DRMFrame3
Public Energy As Long

'game over text
Public WonFrame As Direct3DRMFrame3
Public WonMesh As Direct3DRMMeshBuilder3
Public LostMesh As Direct3DRMMeshBuilder3

'river textures
Public WaterTex(15) As Direct3DRMTexture3
Public WaterCounter As Integer
Public WaterTexCounter As Integer

'logs
Public LogFrame(8) As Direct3DRMFrame3
Public LogMesh As Direct3DRMMeshBuilder3
Public LogPos(8) As D3DVECTOR
Public LogSpeed(8) As Single
Public OnLog As Integer
Public Const NOTONLOG = 12

'the fish
Public FishFrame As Direct3DRMFrame3
Public FishMesh(3) As Direct3DRMMeshBuilder3
Public FishPos As D3DVECTOR
Public FishCount As Integer
Public FishAniFrame As Integer
Public Chasing As Boolean

'the twelve bushes
Public BushFrame(11) As Direct3DRMFrame3
Public BushMesh As Direct3DRMMeshBuilder3
Public BushGone(11) As Boolean

'variable to hold the menu selector position
Public MenuPos As Integer
Public Const MENU_PLAY = 0
Public Const MENU_QUIT = 1

'Following message will be contained in the EXE
Const EXE_Message = "Hey, what do you think you're doing snooping around in this EXE? This is private property! Get out! If you don't leave, you might find the cheat codes! (:O)"

Public Sub SetupDX()
    ChDir App.Path
    
    'set up random number generator
    Randomize
    
    'create Direct3D
    Set D3DRM = DirectX.Direct3DRMCreate()
    Set DirectDraw = DirectX.DirectDraw4Create("")
    Set Clipper = DirectDraw.CreateClipper(0)
    
    'resize form
    If Not Fullscreen Then frmGame.BorderStyle = 1: frmGame.Caption = "Frog3D"
    BorderX = frmGame.Width / Screen.TwipsPerPixelX - frmGame.ScaleWidth
    BorderY = frmGame.Height / Screen.TwipsPerPixelY - frmGame.ScaleHeight
    frmGame.Width = (ScreenX + BorderX) * Screen.TwipsPerPixelX
    frmGame.Height = (ScreenY + BorderY) * Screen.TwipsPerPixelY
    
    'set screen resolution
    If Fullscreen Then DirectDraw.SetDisplayMode ScreenX, ScreenY, 16, 0, DDSDM_DEFAULT
    If ScreenX <> 640 Then frmGame.picLoading.Visible = False: frmGame.lblStatus.Visible = False
    frmGame.Show
    
    'show some progress
    Delay = 600
    frmGame.lblStatus.Caption = "Loading Direct3D"
    DelayGame
    
    'set up the clipper
    Set Clipper = DirectDraw.CreateClipper(0)
    Clipper.SetHWnd frmGame.hWnd
    
    ' init d3drm and main frames
    Set SceneFrame = D3DRM.CreateFrame(Nothing)
    Set LookFrame = D3DRM.CreateFrame(SceneFrame)
    Set CameraFrame = D3DRM.CreateFrame(SceneFrame)
    
    On Error GoTo NoCard
    'uncomment below line to enable the refernce rasterizer (slow)
    'Set device = D3DRM.CreateDeviceFromClipper(clipper, "IID_IDirect3DRGBDevice", frmGame.ScaleWidth, frmGame.ScaleHeight)
    Set Device = D3DRM.CreateDeviceFromClipper(Clipper, "IID_IDirect3DHALDevice", ScreenX, ScreenY)
    
    'set up rendering parameters
    Device.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    'Device.SetDither D_TRUE
    Device.SetShades 1

    'create the viewport
    Set Viewport = D3DRM.CreateViewport(Device, CameraFrame, 0, 0, ScreenX, ScreenY)
    Viewport.SetBack 2000
        
    'set the background color to grey, because it affects credits text
    SceneFrame.SetSceneBackgroundRGB 0.7, 0.7, 0.7
    
    'enable filtering and shading
    If UseFiltering Then Device.SetTextureQuality D3DRMTEXTURE_LINEAR  'smooth out textures
    If UseFlat Then Device.SetQuality D3DRMRENDER_FLAT
    frmGame.picLoading.Width = 50
    
    Delay = 200
    'do the rest
    
    'set up DirectInput
    frmGame.lblStatus.Caption = "Loading DirectInput"
    SetupDInput
    DelayGame
    frmGame.picLoading.Width = 100
    
    'set up DirectSound
    frmGame.lblStatus.Caption = "Loading DirectSound"
    SetupDSound
    DelayGame
    frmGame.picLoading.Width = 175
    
    'set up the lights
    frmGame.lblStatus.Caption = "Setting up Lights"
    SetupLights
    DelayGame
    frmGame.picLoading.Width = 200
    
    'load the world from meshes
    frmGame.lblStatus.Caption = "Loading Meshes"
    LoadWorld
    frmGame.picLoading.Width = 227
    
    'wait a little bit more
    Delay = 500
    DelayGame
    frmGame.picLoading.Visible = False
    
    Exit Sub
NoCard:
    'user has cheap computer
    MsgBox "Frog3D is unable to run because a 3D accelerator card was not found."
    Unload frmGame
End Sub

Sub LoadWorld()
    ChDir App.Path
    
    'create basic world
    Set GroundFrame = D3DRM.CreateFrame(SceneFrame)
    Set GroundMesh = D3DRM.CreateMeshBuilder
    GroundFrame.AddVisual GroundMesh
    GroundMesh.LoadFromFile "ground.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    
    'load logs
    Set LogMesh = D3DRM.CreateMeshBuilder
    For l = 0 To 8
        Set LogFrame(l) = D3DRM.CreateFrame(SceneFrame)
        LogFrame(l).AddVisual LogMesh
    Next
    
    SetupLogs
    
    LogFrame(0).SetPosition Nothing, 0, -0.5, -36.5
    LogFrame(1).SetPosition Nothing, 0, -0.5, -39.6
    LogFrame(2).SetPosition Nothing, 0, -0.5, -42.6
    LogFrame(3).SetPosition Nothing, 0, -0.5, -45.7
    LogFrame(4).SetPosition Nothing, 0, -0.5, -48.7
    LogFrame(5).SetPosition Nothing, 0, -0.5, -51.8
    LogFrame(6).SetPosition Nothing, 0, -0.5, -54.8
    LogFrame(7).SetPosition Nothing, 0, -0.5, -57.9
    LogFrame(8).SetPosition Nothing, 0, -0.5, -60.9
    
    LogMesh.LoadFromFile "log.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    LogMesh.ScaleMesh 0.1, 0.1, 0.1
    OnLog = NOTONLOG
    
    'setup sky
    Set SkyFrame = D3DRM.CreateFrame(SceneFrame)
    Set TempFrame = D3DRM.CreateFrame(SceneFrame)
    Set SkyMesh = D3DRM.CreateMeshBuilder
    
    SkyMesh.LoadFromFile "sky.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    SkyFrame.AddVisual SkyMesh
    Set SkyTexture = D3DRM.LoadTexture("clouds.bmp")
    SkyMesh.SetTexture SkyTexture
    
    SkyMesh.ScaleMesh 1.2, 1.2, 1.2
    
     'make sky even bigger
    SkyMesh.ScaleMesh 3, 3, 3
    
    'create a flat material so sky won't shine
    Dim SkyMat As Direct3DRMMaterial2
    Set SkyMat = D3DRM.CreateMaterial(1)
    SkyMat.SetSpecular 0, 0, 0
    SkyMesh.SetMaterial SkyMat
    
    'create material for cars
    Dim CarMat As Direct3DRMMaterial2
    Set CarMat = D3DRM.CreateMaterial(1)
    CarMat.SetSpecular 0.1, 0.1, 0.1

    'keep ground from shining
    GroundMesh.SetMaterial CarMat
    
    'set up frog
    Set FrogFrame = D3DRM.CreateFrame(SceneFrame)
    For m = 0 To 4
        Set FrogMesh(m) = D3DRM.CreateMeshBuilder
        FrogMesh(m).LoadFromFile "frog" & m + 1 & ".x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        FrogMesh(m).ScaleMesh 0.03, 0.03, 0.03
    Next
    Set SplatMesh = D3DRM.CreateMeshBuilder
    SplatMesh.LoadFromFile "frog1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    SplatMesh.ScaleMesh 0.03, 0.03, 0.03
    FrogFrame.AddVisual FrogMesh(0)
    FrogFrame.SetPosition Nothing, 0, 0.1, 0
    FrogFrame.GetPosition Nothing, FrogPos
    
    'setup piranha fish
    Set FishFrame = D3DRM.CreateFrame(SceneFrame)
    For m = 0 To 3
        Set FishMesh(m) = D3DRM.CreateMeshBuilder
        FishMesh(m).LoadFromFile "fish" & m + 1 & ".x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        FishMesh(m).ScaleMesh 0.08, 0.08, 0.08
    Next
    FishFrame.AddVisual FishMesh(0)
    FishFrame.SetPosition Nothing, 0, -1, -45
    FishFrame.GetPosition Nothing, FishPos

    'load river
    Set WaterFrame = D3DRM.CreateFrame(SceneFrame)
    Set WaterMesh = D3DRM.CreateMeshBuilder
    WaterFrame.AddVisual WaterMesh
    WaterMesh.LoadFromFile "water.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    
    'laod river textures
    For t = 0 To 15
        Set WaterTex(t) = D3DRM.LoadTexture("caust" & t * 2 & ".bmp")
    Next
    
    'setup apple
    Set AppleFrame = D3DRM.CreateFrame(SceneFrame)
    Set AppleMesh = D3DRM.CreateMeshBuilder
    AppleFrame.AddVisual AppleMesh
    AppleMesh.LoadFromFile "apple.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    AppleMesh.ScaleMesh 0.02, 0.02, 0.02
    AppleFrame.SetPosition Nothing, 4.5, 0, -21.5
    
    'load game over meshes
    Set WonFrame = D3DRM.CreateFrame(CameraFrame)
    Set WonMesh = D3DRM.CreateMeshBuilder
    Set LostMesh = D3DRM.CreateMeshBuilder
    'WonFrame.AddVisual WonMesh
    WonMesh.LoadFromFile "youwon.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    WonMesh.ScaleMesh 0.02, 0.02, 0.02
    LostMesh.LoadFromFile "gameover.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    LostMesh.ScaleMesh 0.02, 0.02, 0.02
    
    'setup bee
    Set BeeFrame = D3DRM.CreateFrame(SceneFrame)
    Set AppleBeeFrame = D3DRM.CreateFrame(AppleFrame)
    For m = 0 To 1
        Set BeeMesh(m) = D3DRM.CreateMeshBuilder
        BeeMesh(m).LoadFromFile "bee" & m + 1 & ".x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        BeeMesh(m).ScaleMesh 0.02, 0.02, 0.02
    Next
    'and bee holding apple
    AppleBeeFrame.AddVisual BeeMesh(1)
    AppleBeeFrame.SetPosition AppleBeeFrame, 0, 2, 0
    BeeFrame.AddVisual BeeMesh(1)
    
    BeeFrame.GetPosition Nothing, BeePos
    
    'and fly
    Set FlyFrame = D3DRM.CreateFrame(SceneFrame)
    For m = 0 To 1
        Set FlyMesh(m) = D3DRM.CreateMeshBuilder
        FlyMesh(m).LoadFromFile "fly" & m + 1 & ".x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        FlyMesh(m).ScaleMesh 0.016, 0.016, 0.016
        
    Next
    FlyFrame.AddVisual FlyMesh(1)
    FlyFrame.SetPosition Nothing, 0, 0.3, 0
    
    'set up bushes
    Set BushMesh = D3DRM.CreateMeshBuilder
    BushMesh.LoadFromFile "bush.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    BushMesh.ScaleMesh 0.16, 0.13, 0.15
    For b = 0 To 11
        Set BushFrame(b) = D3DRM.CreateFrame(SceneFrame)
        BushFrame(b).AddVisual BushMesh
    Next
    
    BushFrame(0).SetPosition Nothing, 34, 0, -24.3
    BushFrame(1).SetPosition BushFrame(0), -6, 0, 0
    BushFrame(2).SetPosition BushFrame(1), -6, 0, 0
    BushFrame(3).SetPosition BushFrame(2), -6, 0, 0
    BushFrame(4).SetPosition BushFrame(3), -6, 0, 0
    BushFrame(5).SetPosition BushFrame(4), -6, 0, 0
    BushFrame(6).SetPosition BushFrame(5), -6, 0, 0
    BushFrame(7).SetPosition BushFrame(6), -6, 0, 0
    BushFrame(8).SetPosition BushFrame(7), -6, 0, 0
    BushFrame(9).SetPosition BushFrame(8), -6, 0, 0
    BushFrame(10).SetPosition BushFrame(9), -6, 0, 0
    BushFrame(11).SetPosition BushFrame(10), -6, 0, 0
    
    'select a spot to remove a bush from
    Wack1 = Int(Rnd * 9 + 1)
    
    'same here, except make sure it isn't in the same spot
    Do
        Wack2 = Int(Rnd * 9 + 1)
        DoEvents
    Loop Until Wack1 <> Wack2
    
    BushFrame(Wack1).DeleteVisual BushMesh
    BushGone(Wack1) = True
    BushFrame(Wack2).DeleteVisual BushMesh
    BushGone(Wack2) = True
    
    'set up cars
    Randomize Timer
    
    For m = 0 To 9
        Set CarFrame(m) = D3DRM.CreateFrame(SceneFrame)
        Set CarMesh(m) = D3DRM.CreateMeshBuilder
        CarFrame(m).AddVisual CarMesh(m)
        ChooseCar = Int(Rnd * 3)
        If ChooseCar = 1 Then
            CarMesh(m).LoadFromFile "truck.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        ElseIf ChooseCar = 2 Then
            CarMesh(m).LoadFromFile "car.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Else
            CarMesh(m).LoadFromFile "wagon.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        End If
        CarMesh(m).ScaleMesh 0.07, 0.06, 0.07
        CarFrame(m).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 / 180 * 3.1415927
        CarMesh(m).SetMaterial CarMat
    Next
    For m = 0 To 4
        CarFrame(m).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 180 / 180 * 3.1415927
    Next
    
    SetupCars
    
    'set up menu
    Set PlayFrame = D3DRM.CreateFrame(CameraFrame)
    Set QuitFrame = D3DRM.CreateFrame(CameraFrame)
    Set TitleFrame = D3DRM.CreateFrame(CameraFrame)
    Set MenuFrogFrame = D3DRM.CreateFrame(CameraFrame)
    Set CreditFrame = D3DRM.CreateFrame(CameraFrame)
    
    'energy and score bars
    Set ScoreFrame = D3DRM.CreateFrame(CameraFrame)
    Set ScoreWheelFrame = D3DRM.CreateFrame(ScoreFrame)
    Set EnergyFrame = D3DRM.CreateFrame(CameraFrame)
    Set EnergyWheelFrame = D3DRM.CreateFrame(EnergyFrame)
    Energy = 100
    
    Set PlayMesh = D3DRM.CreateMeshBuilder
    Set QuitMesh = D3DRM.CreateMeshBuilder
    Set TitleMesh = D3DRM.CreateMeshBuilder
    Set MenuFrogMesh = D3DRM.CreateMeshBuilder
    Set CreditMesh = D3DRM.CreateMeshBuilder
    Set ScoreMesh = D3DRM.CreateMeshBuilder
    Set EnergyMesh = D3DRM.CreateMeshBuilder
    Set WheelMesh = D3DRM.CreateMeshBuilder
    
    'create menu, tell it not to delete the score and energy bars
    'because they haven't been shown yet
    CreateMenu True
    
    'load menu meshes
    PlayMesh.LoadFromFile "menuplay.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: PlayMesh.ScaleMesh 0.015, 0.015, 0.02
    QuitMesh.LoadFromFile "menuquit.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: QuitMesh.ScaleMesh 0.015, 0.015, 0.02
    TitleMesh.LoadFromFile "title.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: TitleMesh.ScaleMesh 0.015, 0.015, 0.02
    MenuFrogMesh.LoadFromFile "frog1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: MenuFrogMesh.ScaleMesh 0.005, 0.005, 0.005: MenuFrogFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, -1.1
    CreditMesh.LoadFromFile "credits.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: CreditMesh.ScaleMesh 0.4, 0.4, 0.4
    ScoreMesh.LoadFromFile "scorebar.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: ScoreMesh.ScaleMesh 0.012, 0.0045, 0.013
    EnergyMesh.LoadFromFile "energybar.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: EnergyMesh.ScaleMesh 0.012, 0.0045, 0.013
    WheelMesh.LoadFromFile "wheel.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing: WheelMesh.ScaleMesh 0.01, 0.007, 0.007
    
    'load credits
    Dim CreditTex As Direct3DRMTexture3
    Set CreditTex = D3DRM.LoadTexture("credits.bmp")
    CreditTex.SetDecalTransparency D_TRUE
    CreditMesh.SetTexture CreditTex
    
    'place items in front of screen
    PlayFrame.SetPosition Nothing, 0.114, 0, 2
    TitleFrame.SetPosition Nothing, 0, 0.4, 2
    CreditFrame.SetPosition Nothing, 0, -1.8, 2
    QuitFrame.SetPosition Nothing, 0, -0.3, 2
    MenuFrogFrame.SetPosition Nothing, -0.25, 0, 2
    ScoreFrame.SetPosition Nothing, 0, 0.7, 2
    EnergyFrame.SetPosition Nothing, 0, -0.7, 2
    WonFrame.SetPosition Nothing, 0, 0, 6
    
    'set up camera for the menu
    MFY = 0
    CameraFrame.SetPosition FrogFrame, 0, 10, 0
    CameraFrame.LookAt FrogFrame, Nothing, D3DRMCONSTRAIN_Z
    CameraFrame.SetPosition FrogFrame, 0, 40, 0
    
    CameraView = FREE
    GameMode = MENU
End Sub

Private Sub SetupLights()
    'set up the lights
    Set Light = D3DRM.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.7, 0.7, 0.7)
    Set LightFrame = D3DRM.CreateFrame(SceneFrame)
    LightFrame.SetPosition Nothing, 0, 0, 0
    Light.SetRange 1000!
    LightFrame.AddLight Light
    
    SceneFrame.AddLight D3DRM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.9, 0.9, 0.9)
End Sub

Public Sub SetupCars()
    'put the cars in their respective lanes
    CarFrame(0).SetPosition Nothing, 60, 0, -9: CarSpeed(0) = 0.4
    CarFrame(1).SetPosition Nothing, 35, 0, -9: CarSpeed(1) = 0.4
    
    CarFrame(2).SetPosition Nothing, 60, 0, -2: CarSpeed(2) = 0.32
    CarFrame(3).SetPosition Nothing, 35, 0, -2: CarSpeed(3) = 0.32
    CarFrame(4).SetPosition Nothing, 15, 0, -2: CarSpeed(4) = 0.32
    
    CarFrame(5).SetPosition Nothing, 60, 0, 13: CarSpeed(5) = 0.41
    CarFrame(6).SetPosition Nothing, 30, 0, 13: CarSpeed(6) = 0.41
    
    CarFrame(7).SetPosition Nothing, 60, 0, 6: CarSpeed(7) = 0.33
    CarFrame(8).SetPosition Nothing, 40, 0, 6: CarSpeed(8) = 0.33
    CarFrame(9).SetPosition Nothing, 10, 0, 6: CarSpeed(9) = 0.33
End Sub

Public Sub SetupLogs()
    'place the logs in the river
    Randomize Timer
    LogSpeed(0) = Rnd * 0.05 + 0.07
    For l = 1 To 8
        LogSpeed(l) = LogSpeed(l - 1) + 0.02
        LogFrame(l).SetPosition LogFrame(l), Rnd * 50 - 25, 0, 0
    Next
End Sub

Public Function PosIsClose(frame1 As Direct3DRMFrame3, X As Integer, Z As Integer, Optional MaxDistX As Double, Optional MaxDistZ As Double)
    'check to see if two objects are close or touching
    Dim Vec1 As D3DVECTOR
    Dim Vec2 As D3DVECTOR
    
    frame1.GetPosition Nothing, Vec1
    Vec2.X = X
    Vec2.Z = Z
    
    If MaxDistX <> 0 Then
        'user gave some vars, so use them
        If Vec1.X - MaxDistX < Vec2.X + MaxDistX And Vec1.X + MaxDistX > Vec2.X - MaxDistX Then
            If Vec1.Z - MaxDistZ < Vec2.Z + MaxDistZ And Vec1.Z + MaxDistZ > Vec2.Z - MaxDistZ Then
                PosIsClose = True
            End If
        End If
    Else
        'use default parameters
        If Vec1.X - 5 > Vec2.X + 5 Then
            If Vec1.Z - 5 > Vec2.Z + 5 Then
                PosIsClose = True
            End If
        End If
    End If

End Function

Public Sub DelayGame()
    'Thanks to Fabio Calvi
    StartTick = DirectX.TickCount
    NowTime = DirectX.TickCount
    Do Until NowTime - LastTick > Delay
        DoEvents
        NowTime = DirectX.TickCount
    Loop
    LastTick = NowTime
End Sub

Public Sub CreateMenu(Optional FirstTime As Boolean)
    PlayFrame.AddVisual PlayMesh
    QuitFrame.AddVisual QuitMesh
    TitleFrame.AddVisual TitleMesh
    MenuFrogFrame.AddVisual MenuFrogMesh
    CreditFrame.AddVisual CreditMesh
    If Not FirstTime Then
        ScoreFrame.DeleteVisual ScoreMesh
        EnergyFrame.DeleteVisual EnergyMesh
        ScoreWheelFrame.DeleteVisual WheelMesh
        EnergyWheelFrame.DeleteVisual WheelMesh
    End If
End Sub

Public Sub DeleteMenu()
    PlayFrame.DeleteVisual PlayMesh
    QuitFrame.DeleteVisual QuitMesh
    TitleFrame.DeleteVisual TitleMesh
    MenuFrogFrame.DeleteVisual MenuFrogMesh
    CreditFrame.DeleteVisual CreditMesh
    
    Score = 0
    Energy = 100
    ScoreFrame.AddVisual ScoreMesh
    EnergyFrame.AddVisual EnergyMesh
    
    'reset energy and score bars
    EnergyWheelFrame.SetPosition EnergyFrame, 0.75, 0, 0
    ScoreWheelFrame.SetPosition ScoreFrame, -0.75, 0, 0
    ScoreWheelFrame.AddVisual WheelMesh
    EnergyWheelFrame.AddVisual WheelMesh
End Sub
