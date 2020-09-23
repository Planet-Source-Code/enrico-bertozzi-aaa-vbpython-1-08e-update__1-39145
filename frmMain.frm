VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBPython"
   ClientHeight    =   9975
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   665
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrPwrup 
      Enabled         =   0   'False
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   720
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   0
         Left            =   0
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   1
         Left            =   3000
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   2
         Left            =   6000
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox pp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   3
         Left            =   8880
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Line Line1 
         Index           =   2
         Visible         =   0   'False
         X1              =   584
         X2              =   584
         Y1              =   4
         Y2              =   28
      End
      Begin VB.Label lbL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   30
         Width           =   2535
      End
      Begin VB.Label lbL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   13
         Top             =   30
         Width           =   2535
      End
      Begin VB.Label lbL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   12
         Top             =   30
         Width           =   2535
      End
      Begin VB.Label lbL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9120
         TabIndex        =   11
         Top             =   30
         Width           =   2535
      End
      Begin VB.Label lbS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9120
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.Line Line1 
         Index           =   0
         Visible         =   0   'False
         X1              =   192
         X2              =   192
         Y1              =   4
         Y2              =   28
      End
      Begin VB.Line Line1 
         Index           =   1
         Visible         =   0   'False
         X1              =   392
         X2              =   392
         Y1              =   4
         Y2              =   28
      End
   End
   Begin VB.PictureBox pblack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   210
      Left            =   3120
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox pGame 
      BackColor       =   &H00000000&
      Height          =   9510
      Left            =   0
      ScaleHeight     =   630
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   880
      TabIndex        =   0
      Top             =   480
      Width           =   13260
      Begin VB.Timer tmrGame 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.Label lbPaused 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Paused..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnfNew 
         Caption         =   "&New game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnfPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnfStop 
         Caption         =   "&End game"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnf1 
         Caption         =   "-"
      End
      Begin VB.Menu mnfEsci 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnfOpzioni 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnfHighs 
      Caption         =   "&High scores"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This version lacks of error control routines, modifying something can lead to
'app freeze or crash!

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not InGame Then Exit Sub
If Paused Then Exit Sub
Select Case KeyCode
    Case pKeyUp(1)
    If NumPlayers >= 1 Then 'change direction only if this player is active
    If (FldY(PHX(1), PHY(1)) = -1 Or Not CanChange(1)) And (DontStop = 0 Or DontStop = 1) Then Exit Sub
    FldX(PHX(1), PHY(1)) = 0
    FldY(PHX(1), PHY(1)) = 1
    CanChange(1) = False
    End If
    Case pKeyUp(2)
    If NumPlayers >= 2 Then
    If (FldY(PHX(2), PHY(2)) = -1 + ConPlayer(2) Or Not CanChange(2)) And (DontStop = 0 Or DontStop = 2) Then Exit Sub
    FldX(PHX(2), PHY(2)) = 0 + ConPlayer(2)
    FldY(PHX(2), PHY(2)) = 1 + ConPlayer(2)
    CanChange(2) = False
    End If
    Case pKeyUp(3)
    If NumPlayers >= 3 Then
    If (FldY(PHX(3), PHY(3)) = -1 + ConPlayer(3) Or Not CanChange(3)) And (DontStop = 0 Or DontStop = 3) Then Exit Sub
    FldX(PHX(3), PHY(3)) = 0 + ConPlayer(3)
    FldY(PHX(3), PHY(3)) = 1 + ConPlayer(3)
    CanChange(3) = False
    End If
    Case pKeyUp(4)
    If NumPlayers >= 4 Then
    If (FldY(PHX(4), PHY(4)) = -1 + ConPlayer(4) Or Not CanChange(4)) And (DontStop = 0 Or DontStop = 4) Then Exit Sub
    FldX(PHX(4), PHY(4)) = 0 + ConPlayer(4)
    FldY(PHX(4), PHY(4)) = 1 + ConPlayer(4)
    CanChange(4) = False
    End If
    Case pKeyDn(1)
    If NumPlayers >= 1 Then
    If (FldY(PHX(1), PHY(1)) = 1 Or Not CanChange(1)) And (DontStop = 0 Or DontStop = 1) Then Exit Sub
    FldX(PHX(1), PHY(1)) = 0
    FldY(PHX(1), PHY(1)) = -1
    CanChange(1) = False
    End If
    Case pKeyDn(2)
    If NumPlayers >= 2 Then
    If (FldY(PHX(2), PHY(2)) = 1 + ConPlayer(2) Or Not CanChange(2)) And (DontStop = 0 Or DontStop = 2) Then Exit Sub
    FldX(PHX(2), PHY(2)) = 0 + ConPlayer(2)
    FldY(PHX(2), PHY(2)) = -1 + ConPlayer(2)
    CanChange(2) = False
    End If
    Case pKeyDn(3)
    If NumPlayers >= 3 Then
    If (FldY(PHX(3), PHY(3)) = 1 + ConPlayer(3) Or Not CanChange(3)) And (DontStop = 0 Or DontStop = 3) Then Exit Sub
    FldX(PHX(3), PHY(3)) = 0 + ConPlayer(3)
    FldY(PHX(3), PHY(3)) = -1 + ConPlayer(3)
    CanChange(3) = False
    End If
    Case pKeyDn(4)
    If NumPlayers >= 4 Then
    If (FldY(PHX(1), PHY(4)) = 1 + ConPlayer(4) Or Not CanChange(4)) And (DontStop = 0 Or DontStop = 4) Then Exit Sub
    FldX(PHX(4), PHY(4)) = 0 + ConPlayer(4)
    FldY(PHX(4), PHY(4)) = -1 + ConPlayer(4)
    CanChange(4) = False
    End If
    Case pKeyLe(1)
    If NumPlayers >= 1 Then
    If (FldX(PHX(1), PHY(1)) = -1 Or Not CanChange(1)) And (DontStop = 0 Or DontStop = 1) Then Exit Sub
    FldX(PHX(1), PHY(1)) = 1
    FldY(PHX(1), PHY(1)) = 0
    CanChange(1) = False
    End If
    Case pKeyLe(2)
    If NumPlayers >= 2 Then
    If (FldX(PHX(2), PHY(2)) = -1 + ConPlayer(2) Or Not CanChange(2)) And (DontStop = 0 Or DontStop = 2) Then Exit Sub
    FldX(PHX(2), PHY(2)) = 1 + ConPlayer(2)
    FldY(PHX(2), PHY(2)) = 0 + ConPlayer(2)
    CanChange(2) = False
    End If
    Case pKeyLe(3)
    If NumPlayers >= 3 Then
    If (FldX(PHX(3), PHY(3)) = -1 + ConPlayer(3) Or Not CanChange(3)) And (DontStop = 0 Or DontStop = 3) Then Exit Sub
    FldX(PHX(3), PHY(3)) = 1 + ConPlayer(3)
    FldY(PHX(3), PHY(3)) = 0 + ConPlayer(3)
    CanChange(3) = False
    End If
    Case pKeyLe(4)
    If NumPlayers >= 4 Then
    If (FldX(PHX(4), PHY(4)) = -1 + ConPlayer(4) Or Not CanChange(4)) And (DontStop = 0 Or DontStop = 4) Then Exit Sub
    FldX(PHX(4), PHY(4)) = 1 + ConPlayer(4)
    FldY(PHX(4), PHY(4)) = 0 + ConPlayer(4)
    CanChange(4) = False
    End If
    Case pKeyRi(1)
    If NumPlayers >= 1 Then
    If (FldX(PHX(1), PHY(1)) = 1 Or Not CanChange(1)) And (DontStop = 0 Or DontStop = 1) Then Exit Sub
    FldX(PHX(1), PHY(1)) = -1
    FldY(PHX(1), PHY(1)) = 0
    CanChange(1) = False
    End If
    Case pKeyRi(2)
    If NumPlayers >= 2 Then
    If (FldX(PHX(2), PHY(2)) = 1 + ConPlayer(2) Or Not CanChange(2)) And (DontStop = 0 Or DontStop = 2) Then Exit Sub
    FldX(PHX(2), PHY(2)) = -1 + ConPlayer(2)
    FldY(PHX(2), PHY(2)) = 0 + ConPlayer(2)
    CanChange(2) = False
    End If
    Case pKeyRi(3)
    If NumPlayers >= 3 Then
    If (FldX(PHX(3), PHY(3)) = 1 + ConPlayer(3) Or Not CanChange(3)) And (DontStop = 0 Or DontStop = 3) Then Exit Sub
    FldX(PHX(3), PHY(3)) = -1 + ConPlayer(3)
    FldY(PHX(3), PHY(3)) = 0 + ConPlayer(3)
    CanChange(3) = False
    End If
    Case pKeyRi(4)
    If NumPlayers >= 4 Then
    If (FldX(PHX(4), PHY(4)) = 1 + ConPlayer(4) Or Not CanChange(4)) And (DontStop = 0 Or DontStop = 4) Then Exit Sub
    FldX(PHX(4), PHY(4)) = -1 + ConPlayer(4)
    FldY(PHX(4), PHY(4)) = 0 + ConPlayer(4)
    CanChange(4) = False
    End If
End Select
End Sub

Private Sub Form_Load()
Dim I%
Dim MDetect$

'understanding this game' engine is pretty straighforward:
'there are two arrays, FldX and FldY, each one contains
'information for each block in the field. For example if the block
'is empty, it contains -2, if there is a wall or a power-up, a value
'< -2, and finally if there is a snake, a value > -2.
'for the snakes, each number is the relative location where I
'can find the next piece of the snake:
'
'                 ^
'                 ^
'                 ^<<<<<<<
'

'
'                ^Y-1
'                ^Y-1
'                ^Y-1|<X-1|<X-1|<X-1|<etc...
'

Me.Caption = Version
NumPlayers = 1 'sets default settings
StartingLives = 5
MsgBox KeyInfo, vbInformation 'displays keys assignments

pKeyUp(1) = vbKeyUp
pKeyUp(2) = vbKeyW
pKeyUp(3) = vbKeyNumpad4
pKeyUp(4) = vbKeyI
pKeyDn(1) = vbKeyDown
pKeyDn(2) = vbKeyS
pKeyDn(3) = vbKeyNumpad5
pKeyDn(4) = vbKeyK
pKeyLe(1) = vbKeyLeft
pKeyLe(2) = vbKeyA
pKeyLe(3) = vbKeyNumpad2
pKeyLe(4) = vbKeyJ
pKeyRi(1) = vbKeyRight
pKeyRi(2) = vbKeyD
pKeyRi(3) = vbKeyNumpad8
pKeyRi(4) = vbKeyL

'QX and QY are unit conversion for block>pixels.
'these two FORs sets them to their appropriate value
QSet

FldClear 'clear the game field

For I = 1 To 4 'load snake sprites in each picture box
    pp(I - 1) = LoadPicture(App.Path + "\py" & I & "s.bmp")
Next I

'conplayer is here to distinguish if a value > -2 controls
'a snake or another. It is subtracted each time at the value
'contained in FldX and FldY

ConPlayer(1) = 0
ConPlayer(2) = 3
ConPlayer(3) = 6
ConPlayer(4) = 9

ExpGrowing = True
UseMaps = True

IMDetect = DetectMaps

ReadHS 'read high scores

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized And InGame Then
    tmrGame.Enabled = False
    tmrPwrup.Enabled = False
    Paused = True
    mnfPause.Caption = "&Resume"
    Picture1.Visible = False
    lbPaused.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
WriteHS 'write high scores
End
End Sub

Private Sub mnfEsci_Click()
Unload Me
End Sub

Private Sub mnfHighs_Click()
If Not InGame Then hScores.Show 1, Me Else Beep
End Sub

Private Sub mnfNew_Click()
Dim I%

Randomize Timer
pGame.Cls 'clear the picture box
InGame = True 'set that we are in game
mnfStop.Enabled = True
mnfPause.Enabled = True
mnfNew.Enabled = False
NewSet '>> view for details
tmrGame.Enabled = True

'this timer pop-ups the power-ups objects
tmrPwrup.Enabled = True

For I = 1 To NumPlayers 'reset lifes and score
    Lives(I) = StartingLives
    Score(I) = 0
Next I

LRefresh 'refresh the information labels
Paused = False

For I = 1 To NumPlayers 'set labels to their normal character
    lbS(I - 1).FontBold = False
    lbS(I - 1).ForeColor = RGB(0, 0, 0)
    lbS(I - 1).BackColor = DefaultCol
    lbS(I - 1).FontSize = 8
    lbL(I - 1).FontBold = False
    lbL(I - 1).ForeColor = RGB(0, 0, 0)
    lbL(I - 1).BackColor = DefaultCol
    lbL(I - 1).FontSize = 8
Next I

End Sub

Private Sub mnfOpzioni_Click()
If Not InGame Then frmOptions.Show 1, Me Else Beep
End Sub

Sub ObjDrawAll()

'this sub repaints all the game screen, except for walls, and
'is used in the Paint() event
Dim I%, J%
For I = 1 To 88
    For J = 1 To 63
        If FldX(I, J) <> -2 Then
            Select Case FldX(I, J)
                Case -1, 0, 1
                BitBlt pGame.hDC, QX(I), QY(J), 10, 10, pp(0).hDC, 0, 0, SRCCOPY
                Case 2, 3, 4
                BitBlt pGame.hDC, QX(I), QY(J), 10, 10, pp(1).hDC, 0, 0, SRCCOPY
                Case 5, 6, 7
                BitBlt pGame.hDC, QX(I), QY(J), 10, 10, pp(2).hDC, 0, 0, SRCCOPY
                Case 8, 9, 10
                BitBlt pGame.hDC, QX(I), QY(J), 10, 10, pp(3).hDC, 0, 0, SRCCOPY
                Case -14, -15, -16, -17, -18, -19
                pGame.PaintPicture LoadPicture(App.Path & "\wres" & FldX(I, J) & ".bmp"), QX(I), QY(J)
            End Select
        End If
    Next J
Next I
End Sub

Sub BitDraw(X As Integer, Y As Integer)

'this paint one specified tile in the game screen
Select Case FldX(X, Y)
    Case -1, 0, 1
    BitBlt pGame.hDC, QX(X), QY(Y), 10, 10, pp(0).hDC, 0, 0, SRCCOPY
    Case 2, 3, 4
    BitBlt pGame.hDC, QX(X), QY(Y), 10, 10, pp(1).hDC, 0, 0, SRCCOPY
    Case 5, 6, 7
    BitBlt pGame.hDC, QX(X), QY(Y), 10, 10, pp(2).hDC, 0, 0, SRCCOPY
    Case 8, 9, 10
    BitBlt pGame.hDC, QX(X), QY(Y), 10, 10, pp(3).hDC, 0, 0, SRCCOPY
End Select
End Sub

Private Sub mnfPause_Click()
If Paused Then
    tmrGame.Enabled = True
    tmrPwrup.Enabled = True
    Paused = False
    mnfPause.Caption = "&Pause"
    Picture1.Visible = True
    lbPaused.Visible = False
Else
    tmrGame.Enabled = False
    tmrPwrup.Enabled = False
    Paused = True
    mnfPause.Caption = "&Resume"
    Picture1.Visible = False
    lbPaused.Visible = True
End If
End Sub

Private Sub mnfStop_Click()
Dim I%, J%, PName As String, Place%, ff%
Dim tName As String, tPoints As String

tmrGame.Enabled = False
tmrPwrup.Enabled = False
mnfStop.Enabled = False
mnfNew.Enabled = True
mnfPause.Enabled = False
mnfPause.Caption = "&Pause"
InGame = False
Paused = False
Picture1.Visible = True
lbPaused.Visible = False

'verify whether we made a high score...
For I = 1 To NumPlayers
    Place = IsHighScore(Score(I)) '>> view for details
    If Place > 0 Then 'if we have, request the player name and then
        PName = InputBox("Congratulations player " & I & ", you are " & Place & "° in the high scores. Enter your name:")
        For J = 5 To Place + 1 Step -1
            hSNames(J) = hSNames(J - 1) 'shift down the lower names & points
            hSPoints(J) = hSPoints(J - 1)
        Next J
        hSNames(Place) = PName 'set the appropriate variable
        hSPoints(Place) = CStr(Score(I))
    End If
Next I

FldClear 'clear the game field
pGame.Cls 'clear picturebox
pGame_Paint 'invoke a Paint() event >> view for details (in this case it will paint the game logo)

'EndGame is used as a flag when a match is ended
EndGame = False

'refresh the labels
LRefresh

End Sub

Private Sub pGame_Paint()
If InGame Then 'in case we are ingame, redraw the wall map,
    MapDraw 'and all the objects
    ObjDrawAll
Else 'otherwise, paint the game logo
    pGame.PaintPicture LoadPicture(App.Path & "\logo_final2.jpg"), pGame.Width \ 2 - 400, pGame.Height \ 2 - 300
End If
End Sub

Private Sub tmrGame_Timer()
Static I%, J%, temp%, K%
Dim Crashed As Boolean

    'all these IIFs are used to repeat actions only for enabled players,
    'e.g. if I collected a STOP, other players moves dont need to be computed
    For I = IIf(StopNum > 0, DontStop, 1) To IIf(StopNum > 0, DontStop, NumPlayers)
        PyCrashed(I) = False 'reset variable
    Next I

    Crashed = False
    
    'remember: "PHY(I) = YRedir(tmpHTX, tmpHTY, I)" computes the next Y position (same for X)
    '"FldX(XRedir(PHX(I), PHY(I), I), YRedir(PHX(I), PHY(I), I))" computes the thing in the next X position (same for Y)
    
    For I = IIf(StopNum > 0, DontStop, 1) To IIf(StopNum > 0, DontStop, NumPlayers)
        If Not PyCrashed(I) Then 'here we check if next move, the snake will crash on a wall, or will collect a power-up
            temp = FldX(XRedir(PHX(I), PHY(I), I), YRedir(PHX(I), PHY(I), I))
            
            If (temp >= -13 And temp <= -3) Or (temp >= -1 And temp <= 10) Then
                PyCrashed(I) = True
                Crashed = True
            ElseIf temp >= -19 And temp <= -14 Then
                Collect I, temp
            End If
        End If
    Next I
    
    If Crashed Then Crash 'if flag is set to true, proceed with the SUB
    If EndGame Then 'if the sub sets EndGame, the match is terminated,
        mnfStop_Click 'then invoke a click on the Stop (menu)
        Exit Sub 'then exit sub
    End If
        
    NStep 'compute the next move
    
    If StopNum > 0 Then StopNum = StopNum - 1 'if some snakes are stopped, decrement this
                                            'when we reach zero, other snakes begin to move
For I = 1 To 4
    PyCrashed(I) = False
Next I

Crashed = False
'these controls are made after the move, to verify if two or more snakes' head
'are one over another
For I = 1 To NumPlayers
    For J = I + 1 To NumPlayers
        If PHX(I) = PHX(J) And PHY(I) = PHY(J) Then
            PyCrashed(I) = True
            PyCrashed(J) = True
            Crashed = True
        End If
    Next J
Next I

If Crashed Then Crash
If EndGame Then
    mnfStop_Click
    Exit Sub
End If

'this sets CanChange to true, i.e. the player can now turn
For I = 1 To 4
    CanChange(I) = True
Next I
End Sub

Sub NStep()
Dim I%
Static tmpHTX As Integer, tmpHTY As Integer

For I = IIf(StopNum > 0, DontStop, 1) To IIf(StopNum > 0, DontStop, NumPlayers)
    tmpHTX = PHX(I)
    tmpHTY = PHY(I)
    'set next tile to same value of the head
    FldX(XRedir(PHX(I), PHY(I), I), YRedir(PHX(I), PHY(I), I)) = FldX(PHX(I), PHY(I))
    FldY(XRedir(PHX(I), PHY(I), I), YRedir(PHX(I), PHY(I), I)) = FldY(PHX(I), PHY(I))
    'move the head to the new position
    PHX(I) = XRedir(tmpHTX, tmpHTY, I)
    PHY(I) = YRedir(tmpHTX, tmpHTY, I)
    'draw new head
    BitDraw PHX(I), PHY(I)
    
    'Longer and Shorten are counter that indicate whether a snake
    'must get longer or shorten next times
    If Longer(I) = 0 Then
        tmpHTX = PTX(I) 'if Longer=0, then blank out the tail
        tmpHTY = PTY(I) 'and set new tail position
        PTX(I) = XRedir(tmpHTX, tmpHTY, I)
        PTY(I) = YRedir(tmpHTX, tmpHTY, I)
        FldX(tmpHTX, tmpHTY) = -2
        FldY(tmpHTX, tmpHTY) = -2
        BitBlt pGame.hDC, QX(tmpHTX), QY(tmpHTY), 10, 10, pblack.hDC, 0, 0, SRCCOPY
    Else
        Longer(I) = Longer(I) - 1 'if not, decrement it and
        PyLen(I) = PyLen(I) + 1 'increment snake' lenght
    End If
    
    If Shorten(I) > 0 Then
        tmpHTX = PTX(I) 'if snake must get shorter, an additional
        tmpHTY = PTY(I) 'operation is needed (same as Longer=0)
        PTX(I) = XRedir(tmpHTX, tmpHTY, I)
        PTY(I) = YRedir(tmpHTX, tmpHTY, I)
        FldX(tmpHTX, tmpHTY) = -2
        FldY(tmpHTX, tmpHTY) = -2
        BitBlt pGame.hDC, QX(tmpHTX), QY(tmpHTY), 10, 10, pblack.hDC, 0, 0, SRCCOPY
        Shorten(I) = Shorten(I) - 1 'decrement counter
        PyLen(I) = PyLen(I) - 1 'and snake lenght
    End If
Next I
End Sub

Sub Crash()
Dim I%, Message As String
tmrPwrup.Enabled = False 'stop pwup timer

EndGame = False
For I = 1 To NumPlayers 'if someone has got "out of lifes", set EndGame
    If PyCrashed(I) Then Lives(I) = Lives(I) - 1
    If Lives(I) < 0 Then EndGame = True
Next I

For I = 1 To NumPlayers 'mark red who has lost a life
    If PyCrashed(I) Then
        lbS(I - 1).FontBold = True
        lbS(I - 1).ForeColor = RGB(255, 255, 255)
        lbS(I - 1).BackColor = RGB(255, 0, 0)
        lbS(I - 1).FontSize = 9
        lbL(I - 1).FontBold = True
        lbL(I - 1).ForeColor = RGB(255, 255, 255)
        lbL(I - 1).BackColor = RGB(255, 0, 0)
        lbL(I - 1).FontSize = 9
    End If
Next I
DoEvents
Sleep 1200
If EndGame Then Exit Sub
LRefresh 'refresh the labels
DoEvents
Sleep 1300

For I = 1 To NumPlayers 're-mark normal color all the labels
    lbS(I - 1).FontBold = False
    lbS(I - 1).ForeColor = RGB(0, 0, 0)
    lbS(I - 1).BackColor = DefaultCol
    lbS(I - 1).FontSize = 8
    lbL(I - 1).FontBold = False
    lbL(I - 1).ForeColor = RGB(0, 0, 0)
    lbL(I - 1).BackColor = DefaultCol
    lbL(I - 1).FontSize = 8
Next I

DoEvents

pGame.Cls 'clear picturebox
FldClear 'clear field
NewSet '>> view below
tmrPwrup.Enabled = True 'enable pwup timer

End Sub

Sub NewSet()
Dim X%, Y%
Dim Map As Integer
Dim Result As Boolean

Randomize Timer

'select a valid map
If UseMaps Then
ReMap:
    Do
        Map = Int(Rnd * IMDetect) + 1
    Loop Until MapValid(Map)

    If Not MapSet("map" & Map & ".pm") Then
        MsgBox "An error has occurred using the mapfile: map" & Map & ".pm", vbCritical
        MapValid(Map) = False
        GoTo ReMap
    End If

    'then draw it
    MapDraw
End If

'now, set each player' starting position
If NumPlayers >= 1 Then
    PHX(1) = 9
    PHY(1) = 5
    PTX(1) = 5
    PTY(1) = 5
    PyLen(1) = 5
    Longer(1) = 0
    Shorten(1) = 0
    
    FldX(9, 5) = -1
    FldX(8, 5) = -1
    FldX(7, 5) = -1
    FldX(6, 5) = -1
    FldX(5, 5) = -1
    FldY(5, 5) = 0
    FldY(6, 5) = 0
    FldY(7, 5) = 0
    FldY(8, 5) = 0
    FldY(9, 5) = 0
End If

If NumPlayers >= 2 Then
    PHX(2) = 80
    PHY(2) = 5
    PTX(2) = 84
    PTY(2) = 5
    PyLen(2) = 5
    Longer(2) = 0
    Shorten(2) = 0
    
    FldX(80, 5) = 4
    FldX(81, 5) = 4
    FldX(82, 5) = 4
    FldX(83, 5) = 4
    FldX(84, 5) = 4
    FldY(80, 5) = 3
    FldY(81, 5) = 3
    FldY(82, 5) = 3
    FldY(83, 5) = 3
    FldY(84, 5) = 3
End If

If NumPlayers >= 3 Then
    PHX(3) = 9
    PHY(3) = 59
    PTX(3) = 5
    PTY(3) = 59
    PyLen(3) = 5
    Longer(3) = 0
    Shorten(3) = 0
    
    FldX(9, 59) = 5
    FldX(8, 59) = 5
    FldX(7, 59) = 5
    FldX(6, 59) = 5
    FldX(5, 59) = 5
    FldY(9, 59) = 6
    FldY(8, 59) = 6
    FldY(7, 59) = 6
    FldY(6, 59) = 6
    FldY(5, 59) = 6
    
End If

If NumPlayers >= 4 Then
    PHX(4) = 80
    PHY(4) = 59
    PTX(4) = 84
    PTY(4) = 59
    PyLen(4) = 5
    Longer(4) = 0
    Shorten(4) = 0
    
    FldX(80, 59) = 10
    FldX(81, 59) = 10
    FldX(82, 59) = 10
    FldX(83, 59) = 10
    FldX(84, 59) = 10
    FldY(80, 59) = 9
    FldY(81, 59) = 9
    FldY(82, 59) = 9
    FldY(83, 59) = 9
    FldY(84, 59) = 9
End If

'set pwup timer
tmrPwrup.Interval = Rnd * 8000 + 6000
'draw all the objects and snakes
ObjDrawAll
Do 'randomize next apple' position
    X = Int(Rnd * 87 + 1)
    Y = Int(Rnd * 62 + 1)
Loop Until FldX(X, Y) = -2
'place and draw it
PlaceObj X, Y, fldApple
StopNum = 0
DontStop = 0

End Sub

Sub LRefresh()
Dim I%

'this sub refresh the labels, and hides/displays it according
'to the NumPlayers variable
For I = 1 To NumPlayers
    lbS(I - 1).Visible = True
    lbL(I - 1).Visible = True
    pp(I - 1).Visible = True
    If I <> 4 Then Line1(I - 1).Visible = True
Next I

For I = IIf(InGame, NumPlayers + 1, 1) To 4
    lbS(I - 1).Visible = False
    lbL(I - 1).Visible = False
    pp(I - 1).Visible = False
    If I <> 4 Then Line1(I - 1).Visible = False
Next I

For I = 1 To NumPlayers
    lbL(I - 1).Caption = Lives(I) & " lifes"
    lbS(I - 1).Caption = Score(I) & " points"
Next I
End Sub

Sub PlaceObj(X As Integer, Y As Integer, Obj As Integer)
pGame.PaintPicture LoadPicture(App.Path & "\wres" & Obj & ".bmp"), QX(X), QY(Y)
FldX(X, Y) = Obj
End Sub

Private Sub tmrPwrup_Timer()
Randomize Timer
Dim Obj%, X%, Y%

Select Case Int(Rnd * 9)
    Case 0, 1
    Obj = fldInvert
    Case 2, 3
    Obj = fldStop
    Case 4, 5
    Obj = fldShort
    Case 6, 7
    Obj = fldDouble
    Case 8
    Obj = fldLife
End Select

Do 'look up for a free tile
    X = Int(Rnd * 87 + 1)
    Y = Int(Rnd * 62 + 1)
Loop Until FldX(X, Y) = -2

'place power-up in this tile
PlaceObj X, Y, Obj

're-set interval
tmrPwrup.Interval = Rnd * 8000 + 6000

End Sub

Sub Collect(py As Integer, pwrUp As Integer)
Dim X%, Y%

'collects a powerup
Select Case pwrUp
    Case fldApple
        'when we collect an apple
        If ExpGrowing Then
            'if grow.mode is exponential, tail grows as the snake current lenght
            'if lenght is > 40, grow of 30, because things may get messy!
            Score(py) = (Score(py) + 10 + PyLen(py) * 2) 'add points
            If PyLen(py) >= 40 Then Longer(py) = Longer(py) + 30 Else Longer(py) = Longer(py) + PyLen(py)
            LRefresh 'refresh labels
        Else
            Score(py) = Score(py) + PyLen(py) 'less points for this grow.mode
            Longer(py) = 1 'grow tail of 1 block
            LRefresh 'refresh labels
        End If
        
        Do 'place another apple
            X = Int(Rnd * 87 + 1)
            Y = Int(Rnd * 62 + 1)
        Loop Until FldX(X, Y) = -2
        PlaceObj X, Y, fldApple
        
    Case fldStop
        'stop other snakes as long as the collecting snake is long.
        'if it's too short (< 40), then stop for 40 blocks
        If PyLen(py) > 40 Then StopNum = StopNum + PyLen(py) Else StopNum = StopNum + 40
        DontStop = py 'dont stop the collecting python
    Case fldShort
        'shorten the python of half its size. If it's too short,
        'don't do anyting
        If PyLen(py) > 3 Then Shorten(py) = PyLen(py) \ 2
        Score(py) = (Score(py) + 10 + PyLen(py))
        LRefresh
    Case fldDouble
        'double points
        Score(py) = Score(py) * 2
        LRefresh
    Case fldLife
        'add a life
        Lives(py) = Lives(py) + 1
        LRefresh
    Case fldInvert
        'invert the direction
        PyInvert py
End Select
End Sub

Sub PyInvert(Exclude As Integer)
Dim K%, temp1%, TEMP2%, temp3%, temp4%
Dim CX%, CY%, LnX%, LnY%

'may you will not understand it, due to my bad (?) english...!

'       ^
'       ^
'       ^<<<<<<
'
'
'if we simply invert the direction, it will not work:
'       v
'       v
'       v>>>>>>
'
'       ^
'       |
'snake'll split

'then we have to consider a joint block as if was part of the
'next part of snake:

'       v
'       v
'       >>>>>>>
'
'

'WHAT A MESS!
For K = 1 To NumPlayers
    If K = Exclude Then GoTo SkipI 'exclude the collecting snake
    CX = PTX(K) 'CX, CY = current processed block position
    CY = PTY(K)
    LnX = FldX(CX, CY) 'LnX, Y = what direction we are inverting
    LnY = FldY(CX, CY)
    Do
        temp1 = FldX(CX, CY)
        TEMP2 = FldY(CX, CY)
        If FldX(CX, CY) <> LnX Or FldY(CX, CY) <> LnY Then
            temp3 = FldX(CX, CY) 'this happens when we are in a joint block
            temp4 = FldY(CX, CY)
            FldX(CX, CY) = (LnX - ConPlayer(K)) * -1 + ConPlayer(K)
            FldY(CX, CY) = (LnY - ConPlayer(K)) * -1 + ConPlayer(K)
            LnX = temp3 'set next direction
            LnY = temp4
        Else
            'if we are on a straight part, simply invert it
            FldX(CX, CY) = (FldX(CX, CY) - ConPlayer(K)) * -1 + ConPlayer(K)
            FldY(CX, CY) = (FldY(CX, CY) - ConPlayer(K)) * -1 + ConPlayer(K)
        End If
        'set next block to be processed
        CX = XRedir2(CX - (temp1 - ConPlayer(K)))
        CY = YRedir2(CY - (TEMP2 - ConPlayer(K)))
    Loop Until FldX(CX, CY) = -2 'repeat until we get to a blank tile
    temp1 = PHX(K) 'invert head<>tail
    PHX(K) = PTX(K)
    PTX(K) = temp1
    temp1 = PHY(K)
    PHY(K) = PTY(K)
    PTY(K) = temp1
SkipI:
Next K
End Sub

Function IsHighScore(ByVal Points As Integer) As Integer
Dim I%

IsHighScore = 0
For I = 1 To 5
    If Points > CInt(hSPoints(I)) Then
        IsHighScore = I
        Exit For
    End If
Next I

End Function

Function XRedir(X As Integer, Y As Integer, py As Integer) As Integer

'these two functions, are used to calculate the next X/Y position, and rely on
'the FldX/Y variable.

XRedir = X - (FldX(X, Y) - ConPlayer(py))

'map horizontal/vertical sphericity is implemented here: if a snake
'goes beyond the limit, XRedir would be 0 or 89, so here I correct them
If XRedir < 1 Then XRedir = 88
If XRedir > 88 Then XRedir = 1
End Function

Function YRedir(X As Integer, Y As Integer, py As Integer) As Integer
YRedir = Y - (FldY(X, Y) - ConPlayer(py))

If YRedir < 1 Then YRedir = 63
If YRedir > 63 Then YRedir = 1
End Function

Function XRedir2(X As Integer) As Integer
Static temp%

'these two does the same thing of above ones, but they work for
'all values below 1 and above 88, and return only corrected X, not
'next X pos.

temp = X
If X > 88 Then
    Do
        temp = temp - 88
    Loop Until temp >= 1 And temp <= 88
ElseIf X < 1 Then
    Do
        temp = temp + 88
    Loop Until temp >= 1 And temp <= 88
End If

XRedir2 = temp
End Function

Function YRedir2(Y As Integer) As Integer
Static temp%

temp = Y
If Y > 63 Then
    Do
        temp = temp - 63
    Loop Until temp >= 1 And temp <= 63
ElseIf Y < 1 Then
    Do
        temp = temp + 63
    Loop Until temp >= 1 And temp <= 63
End If

YRedir2 = temp
End Function
