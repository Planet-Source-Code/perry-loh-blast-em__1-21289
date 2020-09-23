Attribute VB_Name = "DDraw"
Dim dx As New DirectX7
Dim dd As DirectDraw7
Dim ddsPrimary As DirectDrawSurface7
Dim ddsBack As DirectDrawSurface7
Dim ddsSprite As DirectDrawSurface7
Dim ddsCrossHair As DirectDrawSurface7
Dim rectSprite As RECT, rectClearBuffer As RECT, rectCrosshair As RECT


Public Sub Initialize(prmForm As Form)
    'On Error GoTo ErrRoutine
    
    Dim ddsd As DDSURFACEDESC2
    
    ' Hide Mouse
    Do Until ShowCursor(0) < 0
    Loop
    
    ' Create direct draw object
    Set dd = dx.DirectDrawCreate("")
    
    ' Set Coop level
    Call dd.SetCooperativeLevel(prmForm.hWnd, DDSCL_ALLOWREBOOT Or DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE)
    
    ' Set display mode
    Call dd.SetDisplayMode(SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_BPP, 0, DDSDM_DEFAULT)
    
    ' Set valid surface flags
    ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    
    ' Set surface capabilities and back buffer count
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    
    ' Create primary surface
    Set ddsPrimary = dd.CreateSurface(ddsd)
    
    ' Set back buffer suface capabilities
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    
    ' Create back buffer surface
    Set ddsBack = ddsPrimary.GetAttachedSurface(caps)
    
    ' Set size of clear buffer rect
    rectClearBuffer.Bottom = SCREEN_HEIGHT: rectClearBuffer.Right = SCREEN_WIDTH
    
    ' Set size of crosshair rect
    rectCrosshair.Bottom = 32: rectCrosshair.Right = 32
    

    ' Load tiles
    Call LoadTiles
    
    
    
    ddsBack.SetForeColor (vbBlue)
    
    Exit Sub
    
ErrRoutine:
    MsgBox Err.Number & " " & Err.Description, vbCritical
    
End Sub

Public Sub Terminate(prmForm As Form)
    On Error GoTo ErrRoutine
    
    ' Release objects
    Set ddsPrimary = Nothing
    Set ddsBack = Nothing
    Set ddsDesktop = Nothing
    
    ' Restore previous display settings
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(prmForm.hWnd, DDSCL_NORMAL)

    ' Show mouse
    Do Until ShowCursor(1) > 0
    Loop
    
    Exit Sub
    
ErrRoutine:
    MsgBox Err.Number & " " & Err.Description, vbCritical
End Sub

Public Sub ClearBuffer()
    
    ' Fill with black
    Call ddsBack.BltColorFill(rectClearBuffer, 0)
    
End Sub

Private Sub LoadTiles()
    On Error GoTo ErrRoutine
    
    Dim ddsd As DDSURFACEDESC2
    Dim CKey As DDCOLORKEY
    
    ' Set valid flags
    ddsd.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    
    ' Set surface capabilities
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ' Set width and height of surface
    ddsd.lWidth = TILE_WIDTH * 16
    ddsd.lHeight = TILE_HEIGHT * 5
    
    ' Create surface from file
    Set ddsSprite = dd.CreateSurfaceFromFile(App.Path & "\LittleMan.bmp", ddsd)
    
    ' Set color key
    CKey.high = vbWhite
    CKey.low = vbWhite
    Call ddsSprite.SetColorKey(DDCKEY_SRCBLT, CKey)
    
    ' Create for crosshair surface
    ddsd.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    
    ' Set surface capabilities
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ' Set width and height of surface
    ddsd.lWidth = 32
    ddsd.lHeight = 32
    
    ' Create surface from file
    Set ddsCrossHair = dd.CreateSurfaceFromFile(App.Path & "\crosshair.bmp", ddsd)
    
    ' Set color key
    CKey.high = vbWhite
    CKey.low = vbWhite
    Call ddsCrossHair.SetColorKey(DDCKEY_SRCBLT, CKey)
    
    
    Exit Sub
    
ErrRoutine:
    
    MsgBox Err.Description
End Sub

Public Function LostSurfaces() As Boolean
    '// Check if we should reload our bitmaps or not
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    
    '// Lost bitmaps, restore the surfaces and return 'true'
    DoEvents
    
    If LostSurfaces Then
        dd.RestoreAllSurfaces
        Call LoadTiles
    End If
    
End Function

Public Function ExclusiveMode() As Boolean
    Dim lTestExMode As Long
    
    '// Test if we're still in exclusive mode
    lTestExMode = dd.TestCooperativeLevel
    
    If (lTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
End Function

Public Sub Flip()

    ' Flip the surface
    Call ddsPrimary.Flip(Nothing, DDFLIP_WAIT)
    
End Sub

Public Sub DrawSprites()
    Dim i As Integer
    
    For i = 0 To SpriteMax
        Sprites(i).Move

        ' Set sprite animation based on frame
        If Sprites(i).State <> 10 Then
            With rectSprite
                .Top = TileHeightOffset(Sprites(i).Color) 'Sprites(i).Color * TILE_HEIGHT
                .Left = TileWidthOffset(Sprites(i).AnimFrame) 'AnimOffset(Sprites(i).AnimFrame) * TILE_WIDTH
                .Bottom = .Top + TILE_HEIGHT
                .Right = .Left + TILE_WIDTH
            End With
        Else
            With rectSprite
                .Top = TileHeightOffset(4) 'Sprites(i).Color * TILE_HEIGHT
                .Left = Sprites(i).AnimFrame * TILE_WIDTH 'AnimOffset(Sprites(i).AnimFrame) * TILE_WIDTH
                .Bottom = .Top + TILE_HEIGHT
                .Right = .Left + TILE_WIDTH
            End With
        End If
        
        Call ddsBack.BltFast(Sprites(i).X, Sprites(i).Y, ddsSprite, rectSprite, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Call ddsBack.BltFast(DInput.X, DInput.Y, ddsCrossHair, rectCrosshair, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        'Call CheckCollision(i)

        Call ddsBack.DrawText(0, 0, "FPS = " & lngFPS, False)
        'Call ddsBack.DrawText(0, 16, "X = " & DInput.X, False)
        'Call ddsBack.DrawText(0, 32, "Y = " & DInput.Y, False)
        
    Next i
    
End Sub

Private Sub CheckCollision(index As Integer)
    ' Checks thru all other sprites for collision,
    ' Sets to reverse velocity if there is a collision
    
    Dim dx As Integer, dy As Integer
    
    For i = 0 To SpriteMax
        If i <> index Then
            ' Check amount of overlap from sprite
            dx = Abs(Sprites(index).X - Sprites(i).X)
            dy = Abs(Sprites(index).Y - Sprites(i).Y)
            
            ' Lower width and height of sprite to make a better collision. More realistic
            If dx < TileCollisionOffsetX And dy < TileCollisionOffsetY Then
                
                ' Collision occured. Tell both sprites to reverse velocity
                Sprites(index).ReverseVelocity
                
                Exit Sub
            End If
            
        End If
    Next i
    
End Sub

Public Sub CheckHit(ByVal X As Integer, ByVal Y As Integer)
    Dim rectSpr As RECT, rectMouse As RECT, rectTmp As RECT
    
    With rectMouse
        .Left = X
        .Top = Y
        .Right = .Left + 1
        .Bottom = .Top + 1
    End With
        
    For i = 0 To SpriteMax
        With rectSpr
            .Left = Sprites(i).X
            .Top = Sprites(i).Y
            .Right = .Left + 14
            .Bottom = .Top + 14
        End With
        
        If (IntersectRect(rectTmp, rectSpr, rectMouse)) Then
            Sprites(i).SetDead
            Exit Sub
        End If
        
    Next i
End Sub

Public Function CheckGameOver() As Boolean
    Dim tmp As Boolean
    
    tmp = True
    For i = 0 To SpriteMax
        If Sprites(i).State <> 10 Then
            tmp = False
        End If
    Next i
    
    CheckGameOver = tmp
End Function
