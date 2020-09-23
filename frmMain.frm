VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 02/23/2001 Perry Loh (skeevs@hotmail.com)
'
' A little screen saver like app which I wrote in my spare time. May seem a little messy and literally
' messy on the screen too ;) Well I had wanted it to do more such as the little people "eating" up
' your desktop, but then I'm a little lazy at the moment, so I decided that I might as well show
' this to some people who want something fun. There are still bugs definately as you will see.
' Leave your comments please, I'd like to hear them.

' I'm not trying to promote violence, it's just for fun.

Dim lngTimer As Long
Dim lngFPSTimer As Long
Dim lngKeyTimer As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call DDraw.Terminate(Me)
        Call DInput.Terminate
        End
    End If
End Sub

Private Sub Form_Load()
    Call Globals.Init
    Call DDraw.Initialize(Me)
    Call DInput.Initialize(Me)
    
    Call GameLoop
    
End Sub


Private Sub GameLoop()
    Dim i As Integer

    lngTimer = GetTickCount
    lngFPSTimer = GetTickCount
    
    Do While bolRunning
        
        ' Main code is below
        If (MS_DELAY + lngTimer) <= GetTickCount() Then
            lngTimer = GetTickCount()
            Call LostSurfaces
            
            Call CheckMouse
            If DInput.LButton Then
                If KEY_DELAY + lngKeyTimer <= GetTickCount() Then
                    lngKeyTimer = GetTickCount
                    Call PlaySound(App.Path & "\shotgun.wav", 0, SND_ASYNC)
                    Call CheckHit(DInput.HotSpotRect.Left, DInput.HotSpotRect.Top)
                End If
            End If
            
            Call DDraw.DrawSprites
            Call DDraw.Flip
            Call DDraw.ClearBuffer
            If DDraw.CheckGameOver Then
                Globals.Init
            End If
            
            lngFramesDone = lngFramesDone + 1
        End If
        
        ' FPS code is below
        If (GetTickCount - lngFPSTimer) >= 1000 Then
            lngFPSTimer = GetTickCount
            lngFPS = lngFramesDone
            lngFramesDone = 0
        End If
        
        DoEvents
    Loop
    
End Sub
