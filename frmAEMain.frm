VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   665
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pBlack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   7080
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   18
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   9
      Left            =   11760
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   17
      Tag             =   "12"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   10
      Left            =   12000
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   16
      Tag             =   "13"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   8
      Left            =   11520
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   15
      Tag             =   "11"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   4
      Left            =   10560
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   14
      Tag             =   "7"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   6
      Left            =   11040
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   13
      Tag             =   "9"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   7
      Left            =   11280
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   12
      Tag             =   "9"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   5
      Left            =   10800
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   11
      Tag             =   "7"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   9600
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   10
      Tag             =   "3"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   2
      Left            =   10080
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   9
      Tag             =   "3"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   10320
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Tag             =   "3"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pPart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   9840
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Tag             =   "3"
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pSelected 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9240
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   6
      Top             =   9750
      Width           =   150
   End
   Begin VB.PictureBox pGame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9450
      Left            =   0
      ScaleHeight     =   630
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   880
      TabIndex        =   0
      Top             =   0
      Width           =   13200
   End
   Begin VB.Label lbNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   9480
      Width           =   975
   End
   Begin VB.Line Line6 
      X1              =   872
      X2              =   8
      Y1              =   647
      Y2              =   647
   End
   Begin VB.Label lbClear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11520
      TabIndex        =   22
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Different"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label lbXY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   9720
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   64
      X2              =   64
      Y1              =   664
      Y2              =   648
   End
   Begin VB.Label lbSelect 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   9720
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   632
      X2              =   632
      Y1              =   664
      Y2              =   648
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   536
      X2              =   536
      Y1              =   664
      Y2              =   648
   End
   Begin VB.Label lbStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   816
      X2              =   816
      Y1              =   648
      Y2              =   664
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   288
      X2              =   288
      Y1              =   648
      Y2              =   664
   End
   Begin VB.Label lbExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12240
      TabIndex        =   3
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label lbSave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Same number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label lbOpen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Save:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   9720
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Me.Caption = Version
FileSave = True
FileNew = True
Selected = 0
CurrentMap = 0

QSet

FldClear
AddSnakes
MapDraw

IMDetect = DetectMaps

lbStatus.Caption = "Higher map number: " & IMDetect

pPart(0).PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), 0, 0
pPart(1).PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), 9, 0, -10, 10
pPart(2).PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), 9, 9, -10, -10
pPart(3).PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), 0, 9, 10, -10
pPart(8).PaintPicture LoadPicture(App.Path + "\wres-11.bmp"), 0, 0
pPart(9).PaintPicture LoadPicture(App.Path + "\wres-12.bmp"), 0, 0
pPart(10).PaintPicture LoadPicture(App.Path + "\wres-13.bmp"), 0, 0
pPart(4).PaintPicture LoadPicture(App.Path + "\wres-7.bmp"), 0, 0
pPart(5).PaintPicture LoadPicture(App.Path + "\wres-9.bmp"), 0, 0
pPart(6).PaintPicture LoadPicture(App.Path + "\wres-7.bmp"), 0, 9, 10, -10
pPart(7).PaintPicture LoadPicture(App.Path + "\wres-9.bmp"), 9, 0, -10, 10

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not FileSave Then
    If MsgBox("You haven't saved this map! Exit anyway?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
End If
End Sub

Private Sub Label1_Click()
Dim newNumber%

newNumber = GetSaveLoc

If MsgBox("Map number will be " & newNumber & ". Save?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

RemoveSnakes
SaveMap newNumber
AddSnakes

CurrentMap = newNumber
Me.Caption = Version & " [map " & CurrentMap & "]"
FileSave = True

IMDetect = DetectMaps
End Sub

Private Sub lbClear_Click()
If FileSave = False Then If MsgBox("You haven't saved the map! Proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

FldClear
pGame.Cls
AddSnakes
MapDraw

FileSave = False
End Sub

Private Sub lbExit_Click()
Unload Me
End Sub

Private Sub lbNew_Click()
If FileSave = False Then If MsgBox("You haven't saved the map! Proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

FldClear
pGame.Cls
AddSnakes
MapDraw

FileSave = False
CurrentMap = 0

Me.Caption = Version
End Sub

Private Sub lbOpen_Click()
Dim MapN As Integer, temp As String

If FileSave = False Then If MsgBox("You haven't saved this map! Proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

Do
    temp = InputBox("Enter map number:")
Loop Until IsNumeric(temp) Or temp = ""

If temp = "" Then Exit Sub

MapN = CInt(temp)

If MapN > IMDetect Then
    MsgBox "Selected map doesn't exists!", vbExclamation
    Exit Sub
End If

If Not MapValid(MapN) Then
    MsgBox "Mapfile is invalid or doesn't exists!", vbCritical
    Exit Sub
End If

FileSave = True
FileNew = False
CurrentMap = MapN
Me.Caption = Version & " [map " & CurrentMap & "]"

FldClear
pGame.Cls
If Not MapSet("map" & MapN & ".pm") Then
    MsgBox "Corrupt mapfile: map" & MapN & ".pm!", vbCritical
    MapValid(MapN) = False
    On Error Resume Next
    Kill App.Path & "\map" & MapN & ".pm"
    Me.Caption = Version & " [recovered data]"
    FileSave = False
    FileNew = True
    CurrentMap = 0
End If
AddSnakes
MapDraw
End Sub

Private Sub lbSave_Click()
If FileNew Then
    Beep
    Exit Sub
End If

If MsgBox("Map will be overwritten! Save anyway?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

RemoveSnakes
SaveMap CurrentMap
AddSnakes
FileSave = True

IMDetect = DetectMaps
End Sub

Private Sub pGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I%, XClick%, YClick%

For I = 1 To 88
    If X > QX(I) Then XClick = I
Next I

For I = 1 To 63
    If Y > QY(I) Then YClick = I
Next I

If FldX(XClick, YClick) > -2 Then
    Beep
    Exit Sub
End If

If Button = 1 Then
    If Selected = 0 Then
        Beep
        Exit Sub
    End If

    FileSave = False

    FldX(XClick, YClick) = -1 * Selected
    DrwSpriteBlt pGame, QX(XClick), QY(YClick), pPart(Selected - 3)

ElseIf Button = 2 And FldX(XClick, YClick) <> fldNull Then
    
    FileSave = False
    FldX(XClick, YClick) = fldNull
    BitBlt pGame.hdc, QX(XClick), QY(YClick), 10, 10, pBlack.hdc, 0, 0, vbSrcCopy
    
End If

End Sub

Private Sub pGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I%, XClick%, YClick%

For I = 1 To 88
    If X > QX(I) Then XClick = I
Next I

For I = 1 To 63
    If Y > QY(I) Then YClick = I
Next I

lbXY.Caption = "X=" & XClick & "   Y=" & YClick
End Sub

Private Sub pGame_Paint()
MapDraw
End Sub

Private Sub pPart_Click(Index As Integer)
Selected = Index + 3
DrwSpriteBlt pSelected, 0, 0, pPart(Index)
End Sub
