Attribute VB_Name = "Globals"

' Change the amount of sprites here
Public Const MAX_SPRITES = 60

Public Sprites() As New clsSprite
Public SpriteMax As Integer
Public bolRunning As Boolean
Public lngFPS As Long, lngFramesDone As Long
Public ScreenWidthTile As Integer
Public ScreenHeightTile As Integer
Public TileCollisionOffsetX As Integer
Public TileCollisionOffsetY As Integer
Public AnimOffset(31) As Integer
Public TileHeightOffset(4) As Integer
Public TileWidthOffset(31) As Integer


' Timer constants
Public Const MS_DELAY = 25
Public Const KEY_DELAY = 1000

' Tile constants
Public Const TILE_WIDTH = 20
Public Const TILE_HEIGHT = 20

' Screen constants
Public Const SCREEN_WIDTH = 800
Public Const SCREEN_HEIGHT = 600
Public Const SCREEN_BPP = 16

Public Const MOUSE_SPEED = 1

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_PRINT = &H2A
Public Const VK_SNAPSHOT = &H2C
Public Const SND_ASYNC = &H1

Public Const SND_SYNC = &H0

Public Sub Init()

    ' This function contains all code that initializes all variables etc.
    ' As you can see, there are a lot of arrays that are precalculated so
    ' in the game loop later, we just have to look them up as a reference table
    
    Dim i As Integer
    
    ' Init maximum number of sprites
    SpriteMax = MAX_SPRITES - 1
    ReDim Sprites(SpriteMax)
    
    ' Initializes all sprites
    For i = 0 To SpriteMax
        Sprites(i).Initialize
    Next i
    
    
    ' Precalculate the screen - tile size
    ScreenWidthTile = SCREEN_WIDTH - TILE_WIDTH
    ScreenHeightTile = SCREEN_HEIGHT - TILE_HEIGHT
    
    ' Precalculate tile collision offset(used for collision)
    TileCollisionOffsetX = TILE_WIDTH - 4
    TileCollisionOffsetY = TILE_HEIGHT - 2
    
    ' Precalculates animation clock speed, this controls the speed of the animation
    For i = 0 To 31
        AnimOffset(i) = Int(i / 2)
    Next i
    
    ' Precalculates tile heights for sprite color offsets
    For i = 0 To 4
        TileHeightOffset(i) = TILE_HEIGHT * i
    Next i
    
    ' Precalculates tile widths for sprite frame offsets
    For i = 0 To 31
        TileWidthOffset(i) = AnimOffset(i) * TILE_WIDTH
    Next i
    
    bolRunning = True
    
End Sub

Public Sub CaptureScreen()
    ' Send Key Event
    Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
    DoEvents
    
    ' Save picture into file
    SavePicture Clipboard.GetData(0), App.Path & "\tmp.bmp"

End Sub
