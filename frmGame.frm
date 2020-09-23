VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ControlBox      =   0   'False
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmGame.frx":2CFA
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3150
      Picture         =   "frmGame.frx":3004
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   6270
      Width           =   30
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   2985
      TabIndex        =   1
      Top             =   6750
      Width           =   3795
   End
   Begin VB.Image imgLoading 
      Height          =   7200
      Left            =   0
      Picture         =   "frmGame.frx":398C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------
'-         Frog3D           -
'----------------------------
'-    This game is based    -
'- upon the game by Konami. -
'----------------------------
'-  (c) 2003-2004 GameSoft  -
'----------------------------
'
Dim Cheating As Boolean

Private Sub Form_Load()
    'load settings from file
    ReadConfig
    'setup directx and build world
    SetupDX
    'start menu
    MenuLoop
End Sub

Sub MenuLoop()
    On Error Resume Next
    GameMode = MENU
    
    'reset the cars and logs back to their normal speeds
    SetupCars
    SetupLogs
    
    'once music stops, start it again
    If MusicBuffer.GetStatus <> DSBSTATUS_PLAYING And UseMusic Then PlaySound MusicBuffer, False, True
    'play ambient backgroud noise
    PlaySound AmbientBuffer(0), True, True
    PlaySound AmbientBuffer(1), True, True

    Delay = GameDelay
    FrogCount = 19
    
    'start loop
    Do
        DoEvents
        DelayGame
        CheckCars
        CheckBee
        CheckFly
        CheckLogs
        CheckFish
        CheckCamera
        CheckKeyboard
        
        RenderScene
    Loop
End Sub

Sub GameLoop()
    On Error Resume Next
    GameMode = GAME
    'change to high camera view
    CameraView = SKY
    'change loop interval to given setting
    Delay = GameDelay
    FrogCount = 19
    
    Do
        DoEvents
        DelayGame
        
        CheckCars
        CheckKeyboard
        CheckBee
        CheckFly
        CheckLogs
        CheckFrog
        CheckFish
        CheckCamera
        
        Caption = FrogPos.Z
        RenderScene
    Loop
End Sub

Private Sub RenderScene()
    'clear what has been already drawn
    Viewport.Clear D3DRMCLEAR_ALL
    'draw objects
    Viewport.Render SceneFrame
    'reflect changes
    Device.Update
End Sub

Private Sub CheckCars()
    'first lane of cars
    For c = 0 To 4
        CarFrame(c).SetPosition CarFrame(c), 0, 0, -CarSpeed(c)
        CarFrame(c).GetPosition Nothing, CarPos(c)
        If CarPos(c).X > 60 Then CarFrame(c).SetPosition Nothing, -60, 0, CarPos(c).Z
    Next
    'second lane
    For c = 5 To 9
        CarFrame(c).SetPosition CarFrame(c), 0, 0, -CarSpeed(c)
        CarFrame(c).GetPosition Nothing, CarPos(c)
        If CarPos(c).X < -60 Then CarFrame(c).SetPosition Nothing, 60, 0, CarPos(c).Z
    Next

    For c = 0 To 9
        'car hits frog, frog gets hurt
        If PosIsClose(FrogFrame, Int(CarPos(c).X), Int(CarPos(c).Z), 1.8, 1) And Not Cheating And Not Squishing Then
            FrogHit
        End If
        'if frog is close but not touching play a honking sound
        If PosIsClose(FrogFrame, Int(CarPos(c).X), Int(CarPos(c).Z), 4, 1) And Not Squishing Then
            HonkSound = Int(Rnd * 10 + 1)
            If HonkSound = 1 Then
                PlaySound HonkBuffer(0), True, False
            Else
                PlaySound HonkBuffer(1), True, False
            End If
        End If
    Next
End Sub

Private Sub CheckBee()
    Static BeeCount As Double
    Static AppleBeeCount As Double
    
    'make bees look at imaginary spot above frog
    TempFrame.SetPosition Nothing, FrogPos.X, 0.5, FrogPos.Z
    BeeFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    AppleBeeFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    'move bee forward
    BeeFrame.SetPosition BeeFrame, 0, 0, 0.04
    BeeFrame.GetPosition Nothing, BeePos
    'move bee up and down on a sine wave
    BeeFrame.SetPosition Nothing, BeePos.X, Sin(CamAngle / 1.6) / 4 + 0.4, BeePos.Z
    
    'manage the bee animation frames
    BeeCount = BeeCount + 0.9
    If BeeCount > 1 Then
        BeeCount = 0
        If Not BeeRaised Then
            BeeFrame.DeleteVisual BeeMesh(1)
            BeeFrame.AddVisual BeeMesh(0)
        Else
            BeeFrame.DeleteVisual BeeMesh(0)
            BeeFrame.AddVisual BeeMesh(1)
        End If
        BeeRaised = Not (BeeRaised)
    End If
    
    'if bee stings frog then frog must lose energy
    If PosIsClose(BeeFrame, Int(FrogPos.X), Int(FrogPos.Z), 0.4, 0.4) And GameMode = GAME And Energy > 0 And Not Cheating And Not Chasing Then
        BeeFrame.SetPosition Nothing, Rnd * 30 - 15, 0.5, -15
        If OnLog <> NOTONLOG Then
            FrogHit True
        Else
            FrogHit
        End If
    End If
    
    'same as above
    AppleBeeCount = AppleBeeCount + 0.9
    If AppleBeeCount > 1 Then
        AppleBeeCount = 0
        If Not AppleBeeRaised Then
            AppleBeeFrame.DeleteVisual BeeMesh(1)
            AppleBeeFrame.AddVisual BeeMesh(0)
        Else
            AppleBeeFrame.DeleteVisual BeeMesh(0)
            AppleBeeFrame.AddVisual BeeMesh(1)
        End If
        AppleBeeRaised = Not (AppleBeeRaised)
    End If

End Sub

Private Sub CheckFly()
    Static Under As Boolean
    Static UnderCount As Single
    Static FlyCount As Double
    Static DirSwitch
    Static MoveCount As Double
    Static NewX As Single
    Static NewZ As Single
    
    MoveCount = MoveCount + 0.5
    If MoveCount > 200 Or MoveCount = 0.5 Then
        DirSwitch = Not (DirSwitch)
        If DirSwitch Then
            'point fly in random direction
            FlyFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, Rnd * 3.14 - 1.07
        Else
            If Under Then
                TempFrame.SetPosition Nothing, 0, -100.3, -23
            Else
                TempFrame.SetPosition Nothing, 0, 0.3, -23
            End If
            'point fly towards imaginary spot
            FlyFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
        End If
        MoveCount = 1
    End If
    
    'if fly is "under", then it is simply 100 y coords below game
    'level, and continues to fly there!
    
    If Under Then
        UnderCount = UnderCount + 1
        If UnderCount > 800 Then
            Under = False
            UnderCount = 0
            'hide fly by putting it below ground
            FlyFrame.SetPosition FlyFrame, 0, 100, -23
        End If
    End If
    
    'move fly forward
    FlyFrame.SetPosition FlyFrame, 0, 0, 0.08
    FlyFrame.GetPosition Nothing, FlyPos
    
    'make fly move up and down on sine wave
    If Under Then
        FlyFrame.SetPosition Nothing, FlyPos.X, Sin(CamAngle / 1.6) / 6 - 100.3, FlyPos.Z
    Else
        FlyFrame.SetPosition Nothing, FlyPos.X, Sin(CamAngle / 1.6) / 6 + 0.3, FlyPos.Z
    End If
    
    'manage the fly animation frames
    FlyCount = FlyCount + 0.9
    If FlyCount > 1 Then
        FlyCount = 0
        If Not FlyRaised Then
            FlyFrame.DeleteVisual FlyMesh(1)
            FlyFrame.AddVisual FlyMesh(0)
        Else
            FlyFrame.DeleteVisual FlyMesh(0)
            FlyFrame.AddVisual FlyMesh(1)
        End If
        FlyRaised = Not (FlyRaised)
    End If
    
    'frog eats fly
    If PosIsClose(FlyFrame, Int(FrogPos.X), Int(FrogPos.Z), 0.4, 0.4) And GameMode = GAME And FlyPos.y > -2 And Not Under Then
        PlaySound CroakBuffer, False, False
        If Energy < 100 Then Energy = Energy + 10: EnergyWheelFrame.SetPosition EnergyWheelFrame, 0.16, 0, 0
        FlyFrame.SetPosition FlyFrame, 0, -100, -23
        FlyFrame.GetPosition Nothing, FlyPos
        Under = True
    End If
End Sub

Private Sub CheckFrog()
    If FrogCount > 0 Then
        'increment animation counter
        FrogCount = FrogCount + 1
    End If
    If FrogCount > 0 And FrogCount < 20 Then
        If FrogCount < 3 Then
            Select Case FrogDir
            Case DIRFRWD
                FrogFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 0.001 / 180 * 3.14
            Case DIRBACK
                FrogFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 180 / 180 * 3.14
            Case DIRLEFT
                FrogFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, -90 / 180 * 3.14
            Case DIRRGHT
                FrogFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, -270 / 180 * 3.14
            End Select
            OnLog = NOTONLOG
        End If

        FrogFrame.SetPosition Nothing, FrogPos.X, (Cos((FrogCount + 8) / 3) / 2) + 0.6, FrogPos.Z
        FrogFrame.SetPosition FrogFrame, 0, 0, -0.17
        
        FrogFrame.GetPosition Nothing, FrogPos
        
        If FrogPos.y < 0.1 Then FrogPos.y = 0.1
        If FrogPos.X < -37.7 Then FrogPos.X = -37.6
        If FrogPos.X > 37.7 Then FrogPos.X = 37.6
        If FrogPos.Z < -70 Then FrogPos.Z = -70
        If FrogPos.Z > 21.6 Then FrogPos.Z = 21.5
        
        'depeding on how far the frog is through the jumping cycle,
        'add and delete meshes to make it look as if it were moving its legs
        If FrogCount = 2 Then
            FrogFrame.DeleteVisual FrogMesh(0)
            FrogFrame.AddVisual FrogMesh(1)
        ElseIf FrogCount = 5 Then
            FrogFrame.DeleteVisual FrogMesh(1)
            FrogFrame.AddVisual FrogMesh(2)
        ElseIf FrogCount = 10 Then
            FrogFrame.DeleteVisual FrogMesh(2)
            FrogFrame.AddVisual FrogMesh(3)
        ElseIf FrogCount = 15 Then
            FrogFrame.DeleteVisual FrogMesh(3)
            FrogFrame.AddVisual FrogMesh(4)
        ElseIf FrogCount = 19 Then
            FrogFrame.DeleteVisual FrogMesh(4)
            FrogFrame.AddVisual FrogMesh(0)
        End If
    End If
    
    Dim FoundLog As Boolean
    If FrogCount > 21 Then
        'frog landed
        FrogCount = 0
        
        'check if frog landed on any of the logs
        For l = 0 To 8
            If PosIsClose(LogFrame(l), Int(FrogPos.X), Int(FrogPos.Z), 1.5, 0.3) Then
                FoundLog = True: OnLog = l: GoTo SkipLog 'found a log, skip other logs
            End If
        Next
SkipLog:
        If Not FoundLog Then
            'frog missed log
            If FrogPos.Z >= -62 And FrogPos.Z <= -35 Then
                'in water
                PlaySound SplashBuffer, False, False
                FrogHit True
            End If
        End If
    End If

    'frog gets apple
    If PosIsClose(AppleFrame, Int(FrogPos.X), Int(FrogPos.Z), 0.5, 0.5) And GameMode = GAME Then
        Score = Score + 10
        If Score = 60 Then YouWon False Else ScoreWheelFrame.SetPosition ScoreWheelFrame, 0.29, 0, 0
        FrogFrame.SetPosition Nothing, 4.5, 0.2, 21.6
        FrogCount = 22
        FrogFrame.GetPosition Nothing, FrogPos
        SpeedUpGame
        PlaySound CroakBuffer, False, False
    End If
    
    'if frog is currently sitting on log, then move frog to match
    'log's position so it looks like frog is still on it
    If OnLog <> NOTONLOG Then
        TempFrame.SetPosition Nothing, FrogPos.X, FrogPos.y, FrogPos.Z
        If (OnLog And 1) = 1 Then
            FrogFrame.SetPosition TempFrame, -LogSpeed(OnLog), 0, 0
        Else
            FrogFrame.SetPosition TempFrame, LogSpeed(OnLog), 0, 0
        End If
        FrogFrame.GetPosition Nothing, FrogPos
        
        If FrogPos.X < -42 Then FrogHit True
        If FrogPos.X > 42 Then FrogHit True
    End If
    
    'detect collisions with the bushes
    For b = 0 To 11
        If PosIsClose(BushFrame(b), Int(FrogPos.X), Int(FrogPos.Z), 1.6, 1.1) And Not BushGone(b) Then FrogHit
    Next
    
    'if in debug mode then display frog's position
    If Cheating Then Caption = FrogPos.X & "  " & FrogPos.Z
End Sub

Private Sub CheckKeyboard(Optional GotoMenu As Boolean)
    'refresh keyboard information
    DInputDevice.GetDeviceStateKeyboard Keyboard
    
    If Keyboard.Key(DIK_ESCAPE) <> 0 And GameMode = GAME Or GotoMenu Then
        'user pressed ESC or game has ended
        FrogFrame.SetPosition Nothing, 0, 0.1, 0: FrogFrame.GetPosition Nothing, FrogPos: CreateMenu: GameMode = MENU
        LookFrame.SetPosition Nothing, 0, 40, 0
        LookFrame.LookAt GroundFrame, Nothing, D3DRMCONSTRAIN_Z

        CameraView = FREE
        MenuLoop
    End If
    
    If GameMode = MENU Then
        If Keyboard.Key(DIK_RETURN) <> 0 Then
            If MenuPos = MENU_PLAY Then DeleteMenu: CameraView = CHASE: FrogFrame.SetPosition Nothing, 4.5, 0.1, 21.6: FrogFrame.GetPosition Nothing, FrogPos: BeeFrame.SetPosition Nothing, Rnd * 30 - 15, 0.5, -15: GameLoop
            If MenuPos = MENU_QUIT Then Unload Me
        End If
        If Keyboard.Key(DIK_DOWN) <> 0 Then
            If MenuPos = MENU_PLAY Then
            MenuPos = MENU_QUIT
            For i = 0 To 20
                PlayFrame.SetPosition PlayFrame, -0.006, 0, 0
                QuitFrame.SetPosition QuitFrame, 0.006, 0, 0
                MFY = MFY - 0.015
                CheckCars
                CheckCamera
                CheckFly
                CheckBee
                DelayGame
                RenderScene
            Next
            End If
        End If
        If Keyboard.Key(DIK_UP) <> 0 Then
            If MenuPos = MENU_QUIT Then
            MenuPos = MENU_PLAY
            For i = 0 To 20
                PlayFrame.SetPosition PlayFrame, 0.006, 0, 0
                QuitFrame.SetPosition QuitFrame, -0.006, 0, 0
                MFY = MFY + 0.015
                CheckCars
                CheckCamera
                CheckFly
                CheckBee
                DelayGame
                RenderScene
            Next
            End If
        End If
        Exit Sub
    End If
    
    'user wants to jump
    If Keyboard.Key(DIK_RIGHT) <> 0 Then
        If FrogCount = 0 Then FrogCount = 1: FrogDir = DIRRGHT: PlaySound JumpBuffer, False, False
    End If
    If Keyboard.Key(DIK_LEFT) <> 0 Then
        If FrogCount = 0 Then FrogCount = 1: FrogDir = DIRLEFT: PlaySound JumpBuffer, False, False
    End If
    
    If Keyboard.Key(DIK_DOWN) <> 0 Then
        If FrogCount = 0 Then FrogCount = 1: FrogDir = DIRBACK: PlaySound JumpBuffer, False, False
    End If
    If Keyboard.Key(DIK_UP) <> 0 Then
        If FrogCount = 0 Then FrogCount = 1: FrogDir = DIRFRWD: PlaySound JumpBuffer, False, False
    End If
    
    'user pressed the cheat keys
    If Keyboard.Key(DIK_RSHIFT) <> 0 And Keyboard.Key(DIK_RSHIFT) <> 0 Then
        Select Case LCase(InputBox("Enter cheat code:", "Enter Code"))
        
        Case "debugmode"
            Cheating = True
        Case "highview"
            CameraView = HIGH
        Case "nomusic"
            MusicBuffer.Stop
        Case "website"
            Shell "start http://www.gamesoft.vze.com"
        Case "speedup"
            SpeedUpGame
            
        End Select
    End If
    
    'Check for camera buttons
    If Keyboard.Key(DIK_INSERT) <> 0 Then CameraView = CHASE
    If Keyboard.Key(DIK_DELETE) <> 0 Then CameraView = INSIDE
    If Keyboard.Key(DIK_HOME) <> 0 Then CameraView = SKY
    If Keyboard.Key(DIK_END) <> 0 Then CameraView = FREE

End Sub

Private Sub CheckCamera()
    Dim RC As D3DVECTOR
    Dim RCU As D3DVECTOR
    If CameraView = CHASE Then LookFrame.SetPosition Nothing, FrogPos.X, 4, FrogPos.Z + 13
    If CameraView = SKY Then LookFrame.SetPosition Nothing, FrogPos.X, 17, FrogPos.Z + 15

    If CameraView = HIGH Then
        LookFrame.SetPosition Nothing, 0, 80, FrogPos.Z
        TempFrame.SetPosition Nothing, 0, 0, FrogPos.Z - 1
        LookFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    End If

    TempFrame.SetPosition Nothing, FrogPos.X, 0, FrogPos.Z
    If GameMode = GAME And CameraView <> HIGH Then LookFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    
    If CameraView = INSIDE Then LookFrame.SetPosition Nothing, FrogPos.X, 3, FrogPos.Z + 2
    
    LookFrame.GetPosition Nothing, CameraPos
    If CameraView <> HIGH Then SkyFrame.SetPosition LookFrame, 0, 0, 0 Else SkyFrame.SetPosition LookFrame, 0, -200, 0
    
    CamAngle = CamAngle + 0.2
    If CamAngle >= 360 Then CamAngle = 0
    
    If GameMode = MENU Then
        'rotate camera on a circle, while looking at center of road
        LookFrame.SetPosition Nothing, Sin(CamAngle / 180 * 3.14) * 20, 2, Cos(CamAngle / 180 * 3.14) * 20
        LookFrame.LookAt FrogFrame, Nothing, D3DRMCONSTRAIN_Z
        CreditFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -0.005
        
        'move title and frog selector back and forth
        TitleFrame.SetPosition CameraFrame, 0, 0.4, Sin(CamAngle * 4 / 180 * 3.14) / 8 + 2
        
        MenuFrogFrame.SetPosition CameraFrame, Sin(CamAngle * 8 / 180 * 3.14) / 20 - 0.3, MFY, 2
    End If
    
    'rotate apple and move it around a circle
    AppleFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 0.1
    AppleFrame.SetPosition Nothing, Sin(CamAngle / 180 * 3.14) * 37, Sin(CamAngle / 2) / 4 + 0.3, Cos(CamAngle / 180 * 3.14) * 2 - 66
    
    'water animation
    WaterCounter = WaterCounter + 1
    If WaterCounter >= 2 Then
        WaterCounter = 0
        
        WaterMesh.SetTexture WaterTex(WaterTexCounter)
        WaterTexCounter = WaterTexCounter + 1
    
        If WaterTexCounter > 15 Then
            WaterTexCounter = 0
        End If
    
    End If
    cameralag = 10
    Dim CamPos As D3DVECTOR
    Dim LookPos As D3DVECTOR
    CameraFrame.GetPosition Nothing, CamPos
    LookFrame.GetPosition Nothing, LookPos
    
    CameraFrame.SetPosition Nothing, CamPos.X + (LookPos.X - CamPos.X) / cameralag, CamPos.y + (LookPos.y - CamPos.y) / cameralag, CamPos.Z + (LookPos.Z - CamPos.Z) / cameralag
    
    Dim TmpVec As D3DVECTOR
    CameraFrame.GetOrientation Nothing, CamPos, TmpVec
    LookFrame.GetOrientation Nothing, LookPos, TmpVec
    
    CameraFrame.SetOrientation Nothing, CamPos.X + (LookPos.X - CamPos.X) / cameralag, CamPos.y + (LookPos.y - CamPos.y) / cameralag, CamPos.Z + (LookPos.Z - CamPos.Z) / cameralag, 0, 0, -1
    
    'Makes the camera not barrel rolled
    CameraFrame.GetOrientation Nothing, RC, RCU
    CameraFrame.SetOrientation SceneFrame, RC.X, RC.y, RC.Z, 0, 1, 0
End Sub

Private Sub CheckLogs()
    'move the logs
    For l = 0 To 8 Step 2
        
        LogFrame(l).SetPosition LogFrame(l), LogSpeed(l), 0, 0
        LogFrame(l).GetPosition Nothing, LogPos(l)
      
        If LogPos(l).X > 41 Then LogFrame(l).SetPosition LogFrame(l), -91, 0, 0
    Next
    
    For l = 1 To 7 Step 2
        
        LogFrame(l).SetPosition LogFrame(l), -LogSpeed(l), 0, 0
        LogFrame(l).GetPosition Nothing, LogPos(l)
        If LogPos(l).X < -45 Then LogFrame(l).SetPosition LogFrame(l), 90, 0, 0
    Next
    
    For l = 0 To 8
        'make logs roll back and forth
        If l <> OnLog Then LogFrame(l).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, Sin(CamAngle / 8) / 32
    Next
    
End Sub

Public Sub CheckFish()
    'make fish look at imaginary spot in front of fish
    TempFrame.SetPosition Nothing, FrogPos.X, -1.5, FrogPos.Z
    FishFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    
    'make fish move, fast if frog is stuck in water
    FishFrame.SetPosition FishFrame, 0, 0, IIf(Chasing, 0.22, 0.08)
    
    'make fish look at frog
    TempFrame.SetPosition Nothing, FrogPos.X, -1, FrogPos.Z
    FishFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
    FishFrame.GetPosition Nothing, FishPos
    
    'keep fish from going out of river
    If FishPos.Z <= -61 Then FishFrame.SetPosition Nothing, FishPos.X, FishPos.y, -61
    If FishPos.Z >= -36 Then FishFrame.SetPosition Nothing, FishPos.X, FishPos.y, -36
    
    FishCount = FishCount + 1
    
    'manage the animation frames
    If FishCount > IIf(Chasing, 2, 4) Then
        
        If FishAniFrame = 0 Then
            FishFrame.DeleteVisual FishMesh(0)
            FishFrame.AddVisual FishMesh(1)
        ElseIf FishAniFrame = 1 Then
            FishFrame.DeleteVisual FishMesh(1)
            FishFrame.AddVisual FishMesh(2)
        ElseIf FishAniFrame = 2 Then
            FishFrame.DeleteVisual FishMesh(2)
            FishFrame.AddVisual FishMesh(3)
        ElseIf FishAniFrame = 3 Then
            FishFrame.DeleteVisual FishMesh(3)
            FishFrame.AddVisual FishMesh(0)
            FishAniFrame = -1
        End If
        
        FishAniFrame = FishAniFrame + 1
        FishCount = 0
        
    End If
End Sub

Private Sub SpeedUpGame()
    'make cars and logs move faster
    For c = 0 To 9
        CarSpeed(c) = CarSpeed(c) + 0.02
    Next
    For l = 0 To 8
        LogSpeed(l) = LogSpeed(l) + 0.002
    Next
End Sub

Private Sub FrogHit(Optional InWater As Boolean)
    'the frog was hit or stung
    
    If InWater Then
        'make fish chase frog
        FrogFrame.SetPosition FrogFrame, 0, -0.7, 0
        Chasing = True
        Do
            DoEvents
            DelayGame
            CheckFish
            CheckCamera
            CheckCars
            CheckLogs
            CheckBee
            RenderScene
            
            If PosIsClose(FrogFrame, Int(FishPos.X), Int(FishPos.Z), 0.5, 0.5) Then Exit Do
        Loop
        Chasing = False
        'set frog in front of river
        FrogFrame.SetPosition Nothing, 4.5, 0.2, -33.4
    ElseIf FrogPos.Z < 15 And FrogPos.Z > -15 Then
        Squishing = True
        FrogFrame.DeleteVisual FrogFrame.GetVisual(0)
        FrogFrame.AddVisual SplatMesh
        PlaySound SplatBuffer, False, False
        For c = 0 To 50
            DoEvents
            DelayGame
            CheckFish
            CheckCamera
            CheckCars
            CheckLogs
            CheckBee
            SplatMesh.ScaleMesh 1.005, 0.9, 1.005
            RenderScene
            
        Next
        Squishing = False
        ChDir App.Path
        SplatMesh.LoadFromFile "frog.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        SplatMesh.ScaleMesh 0.03, 0.03, 0.03
        FrogFrame.DeleteVisual SplatMesh
        FrogFrame.AddVisual FrogMesh(0)
        'frog is on ground, so set him back to the road
        FrogFrame.SetPosition Nothing, 4.5, 0.2, 21.6
    Else
        'frog is on ground, so set him back to the road
        FrogFrame.SetPosition Nothing, 4.5, 0.2, 21.6
        
    End If
    
    FrogFrame.GetPosition Nothing, FrogPos
    FrogPos.y = 0.1: FrogCount = 0
    OnLog = NOTONLOG
    
    'make the frog croak
    PlaySound CroakBuffer, False, False
    
    'move the energy bar
    Energy = Energy - 10
    EnergyWheelFrame.SetPosition EnergyWheelFrame, -0.16, 0, 0
    'if all energy is used up, then game is over
    If Energy <= 0 Then YouWon True
End Sub

Private Sub YouWon(Lost As Boolean)
    'display different text according to if the user won or lost
    If Not Lost Then WonFrame.AddVisual WonMesh Else WonFrame.AddVisual LostMesh
    
    'move the text
    For i = 0 To 280
        DoEvents
        DelayGame
        
        CheckCars
        CheckCamera
        CheckFrog
        CheckBee
        CheckFly
        
        If i < 100 Then WonFrame.SetPosition WonFrame, 0, 0, -0.04
        If i > 240 Then WonFrame.SetPosition WonFrame, 0, -0.06, 0.1
        RenderScene
    Next
    
    'delete the text
    If Not Lost Then WonFrame.DeleteVisual WonMesh Else WonFrame.DeleteVisual LostMesh
    WonFrame.SetPosition LookFrame, 0, 0, 6
    
    'go back to the menu
    CheckKeyboard True
End Sub

Private Sub ReadConfig()
    On Error GoTo NoFile 'if file exists, get setting
                         'else create file with setting
    Open "frog3d.cfg" For Input As #1
        Input #1, TmpF
        Input #1, TmpX
        Input #1, TmpY
        Input #1, TmpJ
        Input #1, TmpB
        
        Input #1, TmpR
        Input #1, TmpT
        
    Close #1
    If TmpF = "TRUE" Then Fullscreen = True
    If TmpJ = "TRUE" Then UseMusic = True
    If TmpB = "TRUE" Then UseFiltering = True
    If TmpR = "TRUE" Then UseFlat = True
    ScreenX = Val(TmpX)
    ScreenY = Val(TmpY)
    GameDelay = Val(TmpT)
    Exit Sub
    
NoFile:
    Close #1
    Open "frog3d.cfg" For Output As #1
        Print #1, "FALSE"
        Print #1, "640"
        Print #1, "480"
        Print #1, "FALSE"
        Print #1, "TRUE"
        Print #1, "FALSE"
        Print #1, "18"
    Close #1
    UseFiltering = True
    UseMusic = True
    ScreenX = 640
    ScreenY = 480
    GameDelay = 18
    
    'let user know
    MsgBox "No config file was found; creating a new one with the default settings"
End Sub

Private Sub Form_Resize()
    'stretch the loading picture to fit the form
    imgLoading.Width = ScaleWidth
    imgLoading.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'delete all the objects
    Set DirectX = Nothing
    Set LookFrame = Nothing
    Set SceneFrame = Nothing
    Set DirectDraw = Nothing
    Set Clipper = Nothing
    Set D3DRM = Nothing
    Set Device = Nothing
    Set Viewport = Nothing
    Set DSound = Nothing
    End
End Sub

