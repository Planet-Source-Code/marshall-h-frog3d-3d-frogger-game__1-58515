Attribute VB_Name = "modSound"
'----------------------------
'-  Sound Code for Frog3D   -
'----------------------------

Public DSound As DirectSound
Public JumpBuffer As DirectSoundBuffer
Public CroakBuffer As DirectSoundBuffer
Public SplashBuffer As DirectSoundBuffer
Public SplatBuffer As DirectSoundBuffer
Public MusicBuffer As DirectSoundBuffer
Public AmbientBuffer(1) As DirectSoundBuffer
Public HonkBuffer(1) As DirectSoundBuffer

'setting structures
Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Sub SetupDSound()
    Set DSound = DirectX.DirectSoundCreate("")
    
    'see if there was a problem
    If Err.Number <> 0 Then
        MsgBox "Unable to continue, error initializing DirectSound."
        'no point in loading sounds if we can't play them
        Exit Sub
    End If
    
    'set ds so that when the game loses focus (ie user switches to another window)
    'then stop playing sounds
    DSound.SetCooperativeLevel frmGame.hWnd, DSSCL_NORMAL
    
    'load sounds
    Set JumpBuffer = CreateSound("jump.wav")
    Set CroakBuffer = CreateSound("croak.wav")
    Set SplashBuffer = CreateSound("splash.wav")
    Set SplatBuffer = CreateSound("splat.wav")
    Set MusicBuffer = CreateSound("menuhigh.wav")
    Set AmbientBuffer(0) = CreateSound("ambient.wav")
    Set AmbientBuffer(1) = CreateSound("ambient2.wav")
    
    Set HonkBuffer(0) = CreateSound("honk1.wav")
    Set HonkBuffer(1) = CreateSound("honk2.wav")
    
    'sound volumes range from -10000 (silent) to 0 (maximum)
    'we are dimming the sounds a bit
    CroakBuffer.SetVolume -500
    MusicBuffer.SetVolume -600
    AmbientBuffer(0).SetVolume -500
    AmbientBuffer(1).SetVolume -500
    HonkBuffer(0).SetVolume -700
    HonkBuffer(1).SetVolume -700
    SplashBuffer.SetVolume -1100
    SplatBuffer.SetVolume -1100
    JumpBuffer.SetVolume -300
    
End Sub

Function CreateSound(filename As String) As DirectSoundBuffer
    DsDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLVOLUME
    
    Set CreateSound = DSound.CreateSoundBufferFromFile(filename, DsDesc, DsWave)
    If Err.Number <> 0 Then
        MsgBox "Unable to find sound file"
        MsgBox Err.Description
        End
    End If
End Function

Sub PlaySound(Sound As DirectSoundBuffer, CloseFirst As Boolean, LoopSound As Boolean)
  
    If CloseFirst Then
        Sound.Stop
        Sound.SetCurrentPosition 0
    End If
    If LoopSound Then
        Sound.Play 1
    Else
        Sound.Play 0
    End If
End Sub

