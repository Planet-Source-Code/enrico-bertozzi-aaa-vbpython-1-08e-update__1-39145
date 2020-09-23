VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lstMaps 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   2760
      List            =   "frmOptions.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   645
      Width           =   1335
   End
   Begin VB.TextBox txGSpeed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   11
      Top             =   1755
      Width           =   1335
   End
   Begin VB.ComboBox lstGrowMode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmOptions.frx":001F
      Left            =   2760
      List            =   "frmOptions.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1020
      Width           =   1335
   End
   Begin VB.OptionButton pn1 
      Caption         =   "&4"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton pn1 
      Caption         =   "&3"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton pn1 
      Caption         =   "&2"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton pn1 
      Caption         =   "&1"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox txLives 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1395
      Width           =   1335
   End
   Begin VB.CommandButton oOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Maps:"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   705
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Game speed (lower=faster):"
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
      Left            =   600
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Growing mode:"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Number of players:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting lifes:"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
pn1(NumPlayers - 1).Value = True
txLives.Text = CStr(StartingLives)
If ExpGrowing Then lstGrowMode.ListIndex = 0 Else lstGrowMode.ListIndex = 1
If UseMaps Then lstMaps.ListIndex = 0 Else lstMaps.ListIndex = 1
txGSpeed.Text = frmMain.tmrGame.Interval
End Sub

Private Sub oOk_Click()
If Not IsNumeric(txLives.Text) Then MsgBox "Starting lifes not specified or invalid value", vbExclamation: Exit Sub
If CInt(txLives.Text) = 0 Then MsgBox "Starting lifes not specified or invalid value", vbExclamation: Exit Sub
If Not IsNumeric(txGSpeed.Text) Then MsgBox "Game speed unspecified or invalid value", vbExclamation: Exit Sub
If CInt(txGSpeed.Text) < 20 Then MsgBox "Game speed unspecified or invalid value", vbExclamation: Exit Sub
If lstGrowMode.ListIndex = 0 Then ExpGrowing = True Else ExpGrowing = False
If lstMaps.ListIndex = 0 Then UseMaps = True Else UseMaps = False
Unload Me
End Sub

Private Sub pn1_Click(Index As Integer)
NumPlayers = Index + 1
End Sub

Private Sub txGSpeed_Change()
If IsNumeric(txGSpeed.Text) Then frmMain.tmrGame.Interval = CInt(txGSpeed.Text)
End Sub

Private Sub txLives_Change()
If IsNumeric(txLives.Text) Then StartingLives = CInt(txLives.Text)
End Sub

Private Sub txLives_GotFocus()
txLives.SelStart = 0
txLives.SelLength = Len(txLives.Text)
End Sub
