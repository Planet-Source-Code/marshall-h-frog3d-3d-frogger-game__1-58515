Attribute VB_Name = "modInput"
'----------------------------
'-  Input Code for Frog3D   -
'----------------------------

'DirectInput
Public DInput As DirectInput
Public DInputDevice As DirectInputDevice
Public Keyboard As DIKEYBOARDSTATE

Public Sub SetupDInput()
    Set DInput = DirectX.DirectInputCreate
    Set DInputDevice = DInput.CreateDevice("GUID_SysKeyboard")
    
    DInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DInputDevice.SetCooperativeLevel frmGame.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DInputDevice.Acquire  'get control of the keyboard!
End Sub

