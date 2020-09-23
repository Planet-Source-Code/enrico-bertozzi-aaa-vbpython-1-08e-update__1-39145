VERSION 5.00
Begin VB.Form hScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Highest Scores"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label ls 
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1770
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Highest Scores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lp 
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
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   8
      Top             =   2130
      Width           =   1095
   End
   Begin VB.Label lp 
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
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   7
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label lp 
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
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Label lp 
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
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   5
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Label lp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   690
      Width           =   1095
   End
   Begin VB.Label ls 
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
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   2130
      Width           =   3255
   End
   Begin VB.Label ls 
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1410
      Width           =   3255
   End
   Begin VB.Label ls 
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   3255
   End
   Begin VB.Label ls 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   690
      Width           =   3255
   End
   Begin VB.Line Line7 
      X1              =   304
      X2              =   8
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line6 
      X1              =   304
      X2              =   8
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line5 
      X1              =   304
      X2              =   8
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line4 
      X1              =   304
      X2              =   8
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line3 
      X1              =   304
      X2              =   8
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line2 
      X1              =   304
      X2              =   8
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line Line1 
      X1              =   224
      X2              =   224
      Y1              =   40
      Y2              =   160
   End
End
Attribute VB_Name = "hScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim I%

For I = 1 To 5
    ls(I - 1).Caption = I & ") " & hSNames(I)
    lp(I - 1).Caption = hSPoints(I)
Next I

End Sub
