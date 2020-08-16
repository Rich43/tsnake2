VERSION 5.00
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Snake 2.0"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox GemMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox GemSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox EndSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox EndMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox HeadMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox HeadSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox WallSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox WallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox FoodMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox FoodSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   3360
      Top             =   1560
   End
   Begin VB.PictureBox TailMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox TailSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2880
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LeftPos As Double
Dim TopPos As Double

Dim FoodLeft As Integer
Dim FoodTop As Integer

Dim Direction As String
Dim TailLength As Integer

Dim Speed As Double

Dim OldLeft As Integer
Dim OldTop As Integer

Dim NewTop As Integer
Dim Newleft As Integer

Dim GemLeft As Integer
Dim GemTop As Integer
Dim GemYN As Boolean
Dim GemRnd As Integer

Dim FoodHasMoved As Boolean

Dim Map As Integer
Dim MapContent() As String

Private Type SnakePos
    Left As Double
    Top As Double
End Type

Private Type WallPos
    Left As Integer
    Top As Integer
End Type

Private Type StrtPos
    Dir As Integer
    Left As Integer
    Top As Integer
End Type

Dim WallCount As Integer
Dim DefaultMapDirection As String

Dim Level As Integer
Dim Score As Long
Dim Steps As Double 'Double
Dim HitsLeft As Integer
Dim LevelScore As Long
Dim TailParts As Double
Dim Lives As Integer

Dim MsgCount As Integer
Dim MsgCountPrev As Integer

Dim SPos() As SnakePos
Dim WPos() As SnakePos
Dim StartPos As StrtPos

Dim MSplit() As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        Case vbKeyLeft
            If Direction = "Right" Then
            Else
                Direction = "Left"
            End If
            
        Case vbKeyRight
            If Direction = "Left" Then
            Else
                Direction = "Right"
            End If
            
        Case vbKeyUp
            If Direction = "Down" Then
            Else
                Direction = "Up"
            End If

        Case vbKeyDown
            If Direction = "Up" Then
            Else
                Direction = "Down"
            End If

        Case vbKeyA
            If Direction = "Right" Then
            Else
                Direction = "Left"
            End If
            
        Case vbKeyD
            If Direction = "Left" Then
            Else
                Direction = "Right"
            End If
            
        Case vbKeyW
            If Direction = "Down" Then
            Else
                Direction = "Up"
            End If

        Case vbKeyS
            If Direction = "Up" Then
            Else
                Direction = "Down"
            End If

End Select
    
    If KeyCode = vbKeyP Then Pause
    If KeyCode = vbKeyPause Then Pause
    
    If KeyCode = vbKeyEscape Then Call Form_QueryUnload(1, 0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    
        If MenuQuit = True Then GoTo ExitSub
        
    QuestionForm.Show
    
    Exit Sub
    
ExitSub:
    Unload Me
    End
    
End Sub

Private Sub Timer1_Timer()

Select Case Direction

    Case "Right"
        LeftPos = LeftPos + Speed
        
        SPos(0).Left = LeftPos
        SPos(0).Top = TopPos
        
    Case "Left"
        LeftPos = LeftPos - Speed
        
        SPos(0).Left = LeftPos
        SPos(0).Top = TopPos
           
    Case "Up"
        TopPos = TopPos - Speed
        
        SPos(0).Top = TopPos
        SPos(0).Left = LeftPos
            
    Case "Down"
        TopPos = TopPos + Speed
        
        SPos(0).Top = TopPos
        SPos(0).Left = LeftPos
                
End Select
    
    Steps = Steps + 1

    MovePic
    
    CollisionDetection

End Sub

Function MovePic()

MainForm.Cls
    
   For p = TailLength To 1 Step -1
   
        If Not p = 1 Then
            
            BitBlt MainForm.hDC, SPos(p - 1).Left, SPos(p - 1).Top, 10, 10, TailMask.hDC, 0, 0, SRCAND
            BitBlt MainForm.hDC, SPos(p - 1).Left, SPos(p - 1).Top, 10, 10, TailSprite.hDC, 0, 0, SRCPAINT
        Else

                If Direction = "Down" Then
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top, 10, 12, HeadMask.hDC, 0, 0, SRCAND
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top, 10, 12, HeadSprite.hDC, 0, 0, SRCPAINT
                End If
                
                If Direction = "Up" Then
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top - 2, 10, 12, HeadMask.hDC, 0, 12, SRCAND
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top - 2, 10, 12, HeadSprite.hDC, 0, 12, SRCPAINT
                End If

                If Direction = "Left" Then
                    BitBlt MainForm.hDC, SPos(0).Left - 2, SPos(0).Top, 12, 10, HeadMask.hDC, 10, 0, SRCAND
                    BitBlt MainForm.hDC, SPos(0).Left - 2, SPos(0).Top, 12, 10, HeadSprite.hDC, 10, 0, SRCPAINT
                End If

                If Direction = "Right" Then
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top, 12, 10, HeadMask.hDC, 10, 10, SRCAND
                    BitBlt MainForm.hDC, SPos(0).Left, SPos(0).Top, 12, 10, HeadSprite.hDC, 10, 10, SRCPAINT
                End If

        End If

            BitBlt MainForm.hDC, SPos(TailLength).Left, SPos(TailLength).Top, 10, 10, EndMask.hDC, 0, 0, SRCAND
            BitBlt MainForm.hDC, SPos(TailLength).Left, SPos(TailLength).Top, 10, 10, EndSprite.hDC, 0, 0, SRCPAINT
            
                SPos(p).Left = SPos(p - 1).Left
                SPos(p).Top = SPos(p - 1).Top
   Next

        BitBlt MainForm.hDC, FoodLeft, FoodTop, 10, 10, FoodMask.hDC, 0, 0, SRCAND
        BitBlt MainForm.hDC, FoodLeft, FoodTop, 10, 10, FoodSprite.hDC, 0, 0, SRCPAINT

    For J = 1 To WallCount

        BitBlt MainForm.hDC, WPos(J).Left, WPos(J).Top, 10, 10, WallMask.hDC, 0, 0, SRCAND
        BitBlt MainForm.hDC, WPos(J).Left, WPos(J).Top, 10, 10, WallSprite.hDC, 0, 0, SRCPAINT

    Next

If GemYN = True Then
            BitBlt MainForm.hDC, GemLeft, GemTop, 9, 9, GemMask.hDC, 0, 0, SRCAND
            BitBlt MainForm.hDC, GemLeft, GemTop, 9, 9, GemSprite.hDC, 0, 0, SRCPAINT
End If

MainForm.Refresh

End Function

Function MoveFood()

    Randomize Timer
    FoodTop = 10 * Int(Rnd * 49)
    
    Randomize Timer
    FoodLeft = 10 * Int(Rnd * 49)

    FoodHasMoved = True
    
    MoveGem
    
End Function

Function MoveGem()

    Randomize Timer
    GemRnd = 1 * Int(Rnd * 20)
    
    If GemRnd = 1 Then
    
        Randomize Timer
        GemLeft = 10 * Int(Rnd * 49)
    
        Randomize Timer
        GemTop = 10 * Int(Rnd * 49)
        
        GemYN = True
        
    Else
    
        GemTop = -40
        GemLeft = -40
        
    End If

End Function

Function CalculateScore()

    Score = Score + (Level * 1000) - (Steps * 2)
    
    LevelScore = LevelScore + (Level * 1000) - (Steps * 2)
    
    If LevelScore <= 0 Then LevelScore = 0
    
    Steps = 0

    HitsLeft = HitsLeft - 1
    
    If HitsLeft = 0 Then
        LevelScore = 0
        Level = Level + 1
        TailParts = 0
        HitsLeft = 10
        Speed = Speed + 0.3
    End If
    
    If Score <= 0 Then Score = 0
    
        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
    
    If Level = 6 Then
    
        If IsSPCM = True Then
            ExitCSPM
            GoTo LeaveSub
        End If
        
        NextmapForm.Show
    End If
    
LeaveSub:

End Function

Function AddTailPart()

    If Level <= 4 Then
        TailLength = TailLength + 4
        TailParts = TailParts + 4
    End If
    
    If Level <= 6 And Level > 4 Then
        TailLength = TailLength + 2
        TailParts = TailParts + 2
    End If
    
    If Level > 6 Then
        TailLength = TailLength + 1
        TailParts = TailParts + 1
    End If
              
        ReDim Preserve SPos(0 To TailLength)

'Bug fix
For K = 0 To TailLength
        If SPos(K).Left = 0 Then SPos(K).Left = -10
Next

End Function

Function Pause()

    Timer1.Enabled = False
    MsgBox "Game Paused", vbInformation, "Paused"
    Timer1.Enabled = True
    
End Function

Function CollisionDetection()


'Snake To Snake Collision
MsgCountPrev = MsgCount

For D = 20 To TailLength

        If SPos(0).Top + 9 >= SPos(D).Top And SPos(0).Top <= SPos(D).Top + 9 And SPos(0).Left + 9 >= SPos(D).Left And SPos(0).Left <= SPos(D).Left + 9 Then
        
            If MsgCount = 0 Then
                ResetSnakeOnSelfHit False
            End If
            
            MsgCount = MsgCount + 1
            
            Exit For
            
        End If
Next

If MsgCount = MsgCountPrev Then
    MsgCount = 0
End If

'Food collision
    If TopPos + 9 >= FoodTop And TopPos <= FoodTop + 9 And LeftPos + 9 >= FoodLeft And LeftPos <= FoodLeft + 9 Then
            MoveFood
            CalculateScore
            AddTailPart
        End If
        
'Gem collision
    If TopPos + 9 >= GemTop And TopPos <= GemTop + 9 And LeftPos + 9 >= GemLeft And LeftPos <= GemLeft + 9 Then
            Lives = Lives + 1
            GemTop = -10
            GemYN = False
            Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
        End If

'Boarder collision
        If TopPos < 0 Then
            TopPos = 490
            Direction = "Up"
        End If
        
        If TopPos > MainForm.ScaleHeight - 10 Then
            TopPos = 0
            Direction = "Down"
        End If
        
        If LeftPos < 0 Then
            LeftPos = 490
            Direction = "Left"
        End If
        
        If LeftPos > MainForm.ScaleWidth - 10 Then
            LeftPos = 0
            Direction = "Right"
        End If
'Food spawns in snake or wall collision
If FoodHasMoved = True Then

    For L = 0 To TailLength
    
        If SPos(L).Top + 9 >= FoodTop And SPos(L).Top <= FoodTop + 9 And SPos(L).Left + 9 >= FoodLeft And SPos(L).Left <= FoodLeft + 9 Then
            MoveFood
        End If
        
    Next
        
    For H = 0 To WallCount
    
        If WPos(H).Top + 9 >= FoodTop And WPos(H).Top <= FoodTop + 9 And WPos(H).Left + 9 >= FoodLeft And WPos(H).Left <= FoodLeft + 9 Then
            MoveFood
        End If
       
    Next
    
'Gem in wall/snake collision
     
        If GemYN = True Then
            For L = 0 To TailLength
    
            If SPos(L).Top + 9 >= GemTop And SPos(L).Top <= GemTop + 9 And SPos(L).Left + 9 >= GemLeft And SPos(L).Left <= GemLeft + 9 Then
                GemLeft = -30
                GemTop = -30
                GemYN = False
            End If
        
        Next
        
         For H = 0 To WallCount
    
            If WPos(H).Top + 9 >= GemTop And WPos(H).Top <= GemTop + 9 And WPos(H).Left + 9 >= GemLeft And WPos(H).Left <= GemLeft + 9 Then
                GemLeft = -30
                GemTop = -30
            End If
        
        Next

        GemYN = False
    End If
    
End If

On Error Resume Next



'Snake to wall collision
For C = 1 To WallCount

    If TopPos + 9 >= WPos(C).Top And TopPos <= WPos(C).Top + 9 And LeftPos + 9 >= WPos(C).Left And LeftPos <= WPos(C).Left + 9 Then
           ResetSnakeOnSelfHit True
     End If
    
Next
            
End Function

Function ResetGame()

    TailLength = 15
    ReDim Preserve SPos(0 To TailLength)
    
    For K = 0 To TailLength
        If SPos(K).Left = 0 Then SPos(K).Left = -10
    Next
    
    SPos(0).Left = -10
    SPos(0).Top = -10
        
    Timer1.Interval = 15
    Timer1.Enabled = True
    
    Speed = 1.4
    Score = 0
    HitsLeft = 10
    Level = 1
    LevelScore = 0
    TailParts = 0
    Lives = 10
    Steps = 0
    
    If IsCustomMap = False Then
    
        Map = 0
        NextMap
            
    Else
    
        If IsSPCM = True Then
            LoadCustomSPMap
        Else
            LoadCustomMap
        End If
        
    End If
    
    Direction = DefaultMapDirection
    LeftPos = StartPos.Left
    TopPos = StartPos.Top
    
    MovePic
    MoveFood

        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
    
End Function

Function ResetSnakeOnSelfHit(WallHit As Boolean)

For G = 0 To TailLength
    SPos(G).Left = -10
    SPos(G).Top = -10
Next

    TailLength = TailLength - TailParts
    ReDim Preserve SPos(0 To TailLength)

    Direction = DefaultMapDirection
    LeftPos = StartPos.Left
    TopPos = StartPos.Top
    
    Score = Score - LevelScore
    
    LevelScore = 0
    TailParts = 0
    HitsLeft = 10
    Lives = Lives - 1
    Steps = 0
    
    If Lives = 0 Then
        GameOver
        GoTo ExitSub
    End If

    If WallHit = False Then
        MsgBox "The snake has hit himself!" & vbCrLf & "The current level will be reset. You have " & Lives & " lives left.", , "Hit!"
    Else
        MsgBox "The snake has hit the wall!" & vbCrLf & "The current level will be reset. You have " & Lives & " lives left.", , "Hit!"
    End If
    
    'Fix This: Score = 0
        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
        
    MoveFood
    
ExitSub:

End Function

Function NextMap()

Map = Map + 1
Dim MapLine As String
Dim WallPos
Dim StartP
Dim Start As String
Dim StartInfo
WallCount = 0
Steps = 0

If Map = 13 Then
    EndSPGame
    GoTo LeaveFunc
End If

        Open App.Path & "\Maps\Level" & Map & ".TSM" For Input As #1
    
            Do Until EOF(1)
            
                Line Input #1, MapLine
                WallPos = Split(MapLine, "|")
                StartP = Split(MapLine, vbCrLf)
                
                WallCount = WallCount + 1
                ReDim Preserve WPos(0 To WallCount)
                
                WPos(WallCount).Top = WallPos(1)
                WPos(WallCount).Left = WallPos(2)

            Loop
        Close #1
                                    
                  Start = StartP(0)
                  StartInfo = Split(StartP(0), "|")
                  
                  StartPos.Dir = StartInfo(1)
                  StartPos.Left = StartInfo(2)
                  StartPos.Top = StartInfo(3)
                  
                  Select Case StartPos.Dir
                  
                        Case 0
                            DefaultMapDirection = "Up"
                        Case 1
                            DefaultMapDirection = "Left"
                        Case 2
                            DefaultMapDirection = "Down"
                        Case 3
                            DefaultMapDirection = "Right"
                            
                  End Select
      
      WallCount = WallCount - 1
      
    TailLength = 15
    ReDim Preserve SPos(0 To TailLength)
    
    For K = 0 To TailLength
        If SPos(K).Left = 0 Then SPos(K).Left = -10
    Next
    
    SPos(0).Left = -10
    SPos(0).Top = -10
        
    Speed = 1.3
    HitsLeft = 10
    Level = 1
    LevelScore = 0
    TailParts = 0
    
    Direction = DefaultMapDirection
    LeftPos = StartPos.Left
    TopPos = StartPos.Top
    
    MovePic
    MoveFood

        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
        
LeaveFunc:
      
End Function
           
Function LoadGFX()

    EndSprite.Picture = LoadPicture("GFX\Snakes\Def_End_Sprite.bmp")
    EndMask.Picture = LoadPicture("GFX\Snakes\Def_End_Mask.bmp")

    TailSprite.Picture = LoadPicture("GFX\Snakes\Def_Tail_Sprite.bmp")
    TailMask.Picture = LoadPicture("GFX\Snakes\Def_Tail_Mask.bmp")
    
    HeadSprite.Picture = LoadPicture("GFX\Snakes\Def_Head_Sprite.bmp")
    HeadMask.Picture = LoadPicture("GFX\Snakes\Def_Head_Mask.bmp")

    FoodSprite.Picture = LoadPicture("GFX\Objects\Def_Food_Sprite.bmp")
    FoodMask.Picture = LoadPicture("GFX\Objects\Def_Food_Mask.bmp")
    
    GemSprite.Picture = LoadPicture("GFX\Objects\Def_Gem_Sprite.bmp")
    GemMask.Picture = LoadPicture("GFX\Objects\Def_Gem_Mask.bmp")
    
    WallSprite.Picture = LoadPicture("GFX\Walls\Def_Wall_Sprite.bmp")
    WallMask.Picture = LoadPicture("GFX\Walls\Def_Wall_Mask.bmp")

    MainForm.Picture = LoadPicture("GFX\Backgrounds\Default.gif")

End Function

Function GameOver()

    Timer1.Enabled = False
    MsgBox "Game Over"
    MainMenu.Show
    MainMenu.Enabled = True
    
End Function

Function FormUnload()

    If MenuQuit = True Then GoTo ExitSub
        
    Timer1.Enabled = False
      
    If YN = "No" Then
        Timer1.Enabled = True
        Exit Function
    End If
    
    If YN = "Yes" Then
        Unload Me
        End
    End If
    
    Timer1.Enabled = True
    
ExitSub:
    Unload Me
    End

End Function

Function LoadCustomMap()
   
Dim MapLine As String
Dim WallPos
Dim StartP
Dim Start As String
Dim StartInfo
WallCount = 0
Steps = 0

If IsFirstCM = True Then
    Lives = 10
    Score = 0
End If

IsFirstCM = False

Timer1.Enabled = True

        CMCounter = CMCounter + 1
            
            MSplit() = Split(CustomMaps, vbCrLf)
                        
                        If CMCounter = UBound(MSplit()) Then
                            ExitMapPack
                            GoTo ExitFunct
                         End If
                            
                        
                Open App.Path & MSplit(CMCounter) For Input As #1
                
                        Do Until EOF(1)
            
                            Line Input #1, MapLine
                            WallPos = Split(MapLine, "|")
                            StartP = Split(MapLine, vbCrLf)
                
                            WallCount = WallCount + 1
                            ReDim Preserve WPos(0 To WallCount)
                
                            WPos(WallCount).Top = WallPos(1)
                            WPos(WallCount).Left = WallPos(2)

                    Loop

                Close #1

                  Start = StartP(0)
                  StartInfo = Split(StartP(0), "|")
                  
                  StartPos.Dir = StartInfo(1)
                  StartPos.Left = StartInfo(2)
                  StartPos.Top = StartInfo(3)
                  
                  Select Case StartPos.Dir
                  
                        Case 0
                            DefaultMapDirection = "Up"
                        Case 1
                            DefaultMapDirection = "Left"
                        Case 2
                            DefaultMapDirection = "Down"
                        Case 3
                            DefaultMapDirection = "Right"
                            
                  End Select
      
      WallCount = WallCount - 1
      
    TailLength = 15
    ReDim Preserve SPos(0 To TailLength)
    
    For K = 0 To TailLength
        If SPos(K).Left = 0 Then SPos(K).Left = -10
    Next
    
    SPos(0).Left = -10
    SPos(0).Top = -10
        
    Speed = 1.3
    HitsLeft = 10
    Level = 1
    LevelScore = 0
    TailParts = 0
    
    Direction = DefaultMapDirection
    LeftPos = StartPos.Left
    TopPos = StartPos.Top
    
    
    'ResetGame
    MovePic
    MoveFood

        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives
ExitFunct:

End Function

Function LoadCustomSPMap()
   
Dim MapLine As String
Dim WallPos
Dim StartP
Dim Start As String
Dim StartInfo
WallCount = 0
Steps = 0

If IsFirstCM = True Then
    Lives = 10
    Score = 0
End If

IsFirstCM = False

Timer1.Enabled = True
                                    
                Open "Custom_Maps\" & SPCMmap & ".TSM" For Input As #1
                
                        Do Until EOF(1)
            
                            Line Input #1, MapLine
                            WallPos = Split(MapLine, "|")
                            StartP = Split(MapLine, vbCrLf)
                
                            WallCount = WallCount + 1
                            ReDim Preserve WPos(0 To WallCount)
                
                            WPos(WallCount).Top = WallPos(1)
                            WPos(WallCount).Left = WallPos(2)

                    Loop

                Close #1

                  Start = StartP(0)
                  StartInfo = Split(StartP(0), "|")
                  
                  StartPos.Dir = StartInfo(1)
                  StartPos.Left = StartInfo(2)
                  StartPos.Top = StartInfo(3)
                  
                  Select Case StartPos.Dir
                  
                        Case 0
                            DefaultMapDirection = "Up"
                        Case 1
                            DefaultMapDirection = "Left"
                        Case 2
                            DefaultMapDirection = "Down"
                        Case 3
                            DefaultMapDirection = "Right"
                            
                  End Select
      
      WallCount = WallCount - 1
      
    TailLength = 15
    ReDim Preserve SPos(0 To TailLength)
    
    For K = 0 To TailLength
        If SPos(K).Left = 0 Then SPos(K).Left = -10
    Next
    
    SPos(0).Left = -10
    SPos(0).Top = -10
        
    Speed = 1.3
    HitsLeft = 10
    Level = 1
    LevelScore = 0
    TailParts = 0
    
    Direction = DefaultMapDirection
    LeftPos = StartPos.Left
    TopPos = StartPos.Top
    
    
    'ResetGame
    MovePic
    MoveFood

        Me.Caption = "Score: " & Score & "  Level: " & Level & "  Hits left: " & HitsLeft & "  Lives: " & Lives

End Function

Function ExitCSPM()
    IsSPCM = False
    MainForm.Timer1.Interval = 0
    MsgBox "You have completed the map!", , "Completed"
    Me.Enabled = False
    MainMenu.Enabled = True
    MainMenu.Show
End Function
Function EndSPGame()
    MainForm.Timer1.Interval = 0
    MsgBox "Congratulations, you have completed all Singleplayer map!!!", , "Congratulations!!!"
    Me.Enabled = False
    MainMenu.Enabled = True
    MainMenu.Show
End Function
Function ExitMapPack()
    MainForm.Timer1.Interval = 0
    MsgBox "Congratulations, you have completed the Map Pack!", , "Congratulations!"
    Me.Enabled = False
    MainMenu.Enabled = True
    MainMenu.Show
End Function

Private Sub Timer2_Timer()
    GemTop = -30
    GemLeft = -30
    Timer2.Enabled = False
End Sub
