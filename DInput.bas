Attribute VB_Name = "DInput"
Dim dx As New DirectX7
Dim di As DirectInput
Public diDev As DirectInputDevice
Public diMState As DIMOUSESTATE
Public X As Integer
Public Y As Integer
Public LButton As Boolean
Public RButton As Boolean
Private XBorder As Integer
Private YBorder As Integer
Public MouseRect As RECT
Public HotSpotRect As RECT

Public Sub Initialize(prmForm As Form)
    ' Create DI object
    Set di = dx.DirectInputCreate()
    
    ' Create DI device which is a mouse
    Set diDev = di.CreateDevice("GUID_SysMouse")
    
    ' Set the data format,coop level and acquire
    Call diDev.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call diDev.SetCooperativeLevel(prmForm.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
    
    diDev.Acquire
    
    XBorder = SCREEN_WIDTH - 32
    YBorder = SCREEN_HEIGHT - 32
    
    ' Set mouse to center of screen
    X = SCREEN_WIDTH / 2
    Y = SCREEN_HEIGHT / 2
End Sub

Public Sub Terminate()
    diDev.Unacquire
    Set diDev = Nothing
    Set di = Nothing
End Sub

Public Sub CheckMouse()
    ' Get state
    Call diDev.GetDeviceStateMouse(DInput.diMState)
    
    ' Acquire if we lost it
    If Err.Number <> 0 Then DInput.diDev.Acquire
    
    ' Exit if we cannot acquire
    If Err.Number <> 0 Then Exit Sub
    
    ' Calculate new position of mouse
    X = X + diMState.X * MOUSE_SPEED
    If X < 0 Then X = 0
    If X > XBorder Then X = XBorder
    
    
    Y = Y + diMState.Y * MOUSE_SPEED
    If Y <= 0 Then Y = 0
    If Y > YBorder Then Y = YBorder
    
    ' Check Left Button
    If diMState.buttons(0) <> 0 Then LButton = True
    If diMState.buttons(0) = 0 Then LButton = False
    
    ' Check Right Button
    If diMState.buttons(1) <> 0 Then RButton = True
    If diMState.buttons(1) = 0 Then RButton = False
    
    
    ' Update the hotspot rect
    With HotSpotRect
        .Top = Y + 16
        .Left = X + 16
        .Right = X + 1
        .Bottom = Y + 1
    End With
    
End Sub


