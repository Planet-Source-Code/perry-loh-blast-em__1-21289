VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Constants

Const TIMER_DELAY = 10000

Const SPEED_RUNNING = 2
Const SPEED_WALKING = 1

Const STATE_IDLE = 0
Const STATE_WALK = 1
Const STATE_RUN = 3
Const STATE_DEAD = 10

Const DIR_NORTH = 0
Const DIR_EAST = 1
Const DIR_SOUTH = 2
Const DIR_WEST = 3

Const CLR_RED = 0
Const CLR_GREEN = 1
Const CLR_BLUE = 2
Const CLR_LBLUE = 3
Const CLR_DEAD = 4

Private m_X As Integer
Private m_Y As Integer
Private m_XVelocity As Integer
Private m_YVelocity As Integer
Private m_Direction As Byte
Private m_AnimFrame As Byte
Private m_State As Byte
Public m_MaxIdle As Integer
Private m_CurrIdle As Integer
Private m_Timer As Long

Private m_Color As Byte


Public Property Get X() As Integer
    X = m_X
End Property

Public Property Get Y() As Integer
    Y = m_Y
End Property

Public Property Get State() As Byte
    State = m_State
End Property

Public Property Get AnimFrame() As Byte
    AnimFrame = m_AnimFrame
End Property

Public Property Get Direction() As Byte
    Direction = m_Direction
End Property

Public Property Get Color() As Byte
    Color = m_Color
End Property

Public Sub Initialize()
    Randomize GetTickCount()
    
    ' Get a random position on the screen
    m_X = Rnd() * SCREEN_WIDTH
    m_Y = Rnd() * SCREEN_HEIGHT
    
    ' Get a random state, either walk or idle
    m_State = STATE_WALK 'Rnd() * 1
    Call SelectState
    
    ' Set velocity of this sprite
    Call SetVelocity
    
    m_AnimFrame = 0
    m_Color = Rnd() * 3
    m_Timer = GetTickCount
    
End Sub

Public Sub Move()
    
    If m_State <> STATE_DEAD Then
        ' If the sprite has been set to stay idle for the number of cycles,
        ' reduce the count every game loop until it reaches 0, once there
        ' set the sprite back to a walking state and set a random direction
        If m_MaxIdle > 0 Then
            m_MaxIdle = m_MaxIdle - 1
            
            If m_MaxIdle = 0 Then
                m_State = STATE_WALK
                Call SetVelocity
                
            End If
        Else
            
            ' Update animation frame, based on direction
            Select Case m_Direction
                
                Case DIR_NORTH
                    If m_AnimFrame >= 7 Then
                        m_AnimFrame = 0
                    Else
                        m_AnimFrame = m_AnimFrame + 1
                    End If
                
                Case DIR_EAST
                    If m_AnimFrame >= 15 Or m_AnimFrame < 8 Then
                        m_AnimFrame = 8
                    Else
                        m_AnimFrame = m_AnimFrame + 1
                    End If
                
                Case DIR_SOUTH
                    If m_AnimFrame >= 23 Or m_AnimFrame < 16 Then
                        m_AnimFrame = 16
                    Else
                        m_AnimFrame = m_AnimFrame + 1
                    End If
                
                Case DIR_WEST
                    If m_AnimFrame >= 31 Or m_AnimFrame < 24 Then
                        m_AnimFrame = 24
                    Else
                        m_AnimFrame = m_AnimFrame + 1
                    End If
                
            End Select
                
            ' Check for X boundary collision
            If m_X < 0 Then
                m_X = SCREEN_WIDTH
            ElseIf m_X > SCREEN_WIDTH Then
                m_X = 0
            End If
            
            ' Check for Y boundary collision
            If m_Y < 0 Then
                m_Y = SCREEN_HEIGHT
            ElseIf m_Y > SCREEN_HEIGHT Then
                m_Y = 0
            End If
            
            ' Set the X and Y based on the directions
            m_X = m_X + m_XVelocity
            m_Y = m_Y + m_YVelocity
    
        End If
        
        If (TIMER_DELAY + m_Timer) <= GetTickCount() Then
            m_Timer = GetTickCount
            
            ' Stay idle or continue
            Call SelectState
        End If
    Else
        If m_AnimFrame < (m_Color * 4 + 1) Then
            m_AnimFrame = m_AnimFrame + 1
        End If
        
'        If m_AnimFrame < 1 Then
'            m_AnimFrame = m_AnimFrame + 1
'        End If
    End If
End Sub

Public Sub SetDead()
    m_State = STATE_DEAD
    
    Select Case m_Color
        Case CLR_RED
            m_AnimFrame = 0
            
        Case CLR_GREEN
            m_AnimFrame = 4
            
        Case CLR_BLUE
            m_AnimFrame = 8
            
        Case CLR_LBLUE
            m_AnimFrame = 12

    End Select
    
    'm_AnimFrame = m_Color * 5
    'm_Color = CLR_DEAD
End Sub

Private Sub SelectState()
    Dim tmp As Integer
    
    ' Get a random number to test for state
    Randomize GetTickCount
    tmp = Rnd() * 10
    
    If tmp > 7 Then
        If m_State <> STATE_IDLE Then
            m_State = STATE_IDLE
            m_XVelocity = 0
            m_YVelocity = 0
            
            ' Get a random idle time cycle
            m_MaxIdle = Rnd() * 200
            m_CurrIdle = 0
        End If
    ElseIf tmp > 3 Then
        If m_State <> STATE_IDLE Then
            m_State = STATE_RUN
            m_MaxIdle = 0
            m_CurrIdle = 0
        End If
    Else
        If m_State <> STATE_IDLE Then
            m_State = STATE_WALK
            m_MaxIdle = 0
            m_CurrIdle = 0
        End If
    End If
End Sub

Public Sub ReverseVelocity()
    If m_State <> STATE_IDLE Then
        ' Reverse velocity of sprite
        m_XVelocity = -(m_XVelocity)
        m_YVelocity = -(m_YVelocity)
        
        ' Update direction
        Select Case m_Direction
            Case DIR_NORTH
                m_Direction = DIR_SOUTH
                
            Case DIR_EAST
                m_Direction = DIR_WEST
            
            Case DIR_SOUTH
                m_Direction = DIR_NORTH
            
            Case DIR_WEST
                m_Direction = DIR_EAST
        End Select
    End If
End Sub

Public Sub SetVelocity()

    ' Set a random direction
    m_Direction = Rnd() * 3
    
    Select Case m_Direction
        Case DIR_NORTH
            m_XVelocity = 0
            m_YVelocity = -(m_State)
        
        Case DIR_EAST
            m_XVelocity = m_State
            m_YVelocity = 0
        
        Case DIR_SOUTH
            m_XVelocity = 0
            m_YVelocity = m_State
        
        Case DIR_WEST
            m_XVelocity = -(m_State)
            m_YVelocity = 0
        
    End Select
    
End Sub
