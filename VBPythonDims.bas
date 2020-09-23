Attribute VB_Name = "VBPythonDims"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global NumPlayers As Integer, StartingLives As Integer
Global InGame As Boolean, Paused As Boolean
Global StopNum As Integer, DontStop As Integer
Global ExpGrowing As Boolean, UseMaps As Boolean
Global EndGame As Boolean

Global hSNames(1 To 5) As String
Global hSPoints(1 To 5) As String

Global IMDetect As Integer

Global Const DefaultCol As Long = -2147483633 'transparent background color (for labels)

Global QX(1 To 88) As Integer
Global QY(1 To 63) As Integer
Global FldX(1 To 88, 1 To 63) As Integer
Global FldY(1 To 88, 1 To 63) As Integer

Global PHY(1 To 4) As Integer
Global PTY(1 To 4) As Integer
Global PHX(1 To 4) As Integer
Global PTX(1 To 4) As Integer
Global Score(1 To 4) As Long
Global Lives(1 To 4) As Integer

Global pKeyUp(1 To 4) As Integer
Global pKeyDn(1 To 4) As Integer
Global pKeyLe(1 To 4) As Integer
Global pKeyRi(1 To 4) As Integer

Global Longer(1 To 4) As Integer
Global Shorten(1 To 4) As Integer
Global PyLen(1 To 4) As Integer

Global PyCrashed(1 To 4) As Boolean
Global CanChange(1 To 4) As Boolean
Global ConPlayer(1 To 4) As Integer

Global Const fldNull As Integer = -2
Global Const fldMRot1 As Integer = -3
Global Const fldMRot2 As Integer = -4
Global Const fldMRot3 As Integer = -5
Global Const fldMRot4 As Integer = -6
Global Const fldMI1 As Integer = -7
Global Const fldMI2 As Integer = -8
Global Const fldMI3 As Integer = -9
Global Const fldMI4 As Integer = -10
Global Const fldMCross As Integer = -11
Global Const fldMH As Integer = -12
Global Const fldMV As Integer = -13
Global Const fldApple As Integer = -14
Global Const fldInvert As Integer = -15
Global Const fldDouble As Integer = -16
Global Const fldShort As Integer = -17
Global Const fldLife As Integer = -18
Global Const fldStop As Integer = -19

Global Const KeyInfo As String = "Game keys:" & vbCrLf & "in order: up, down, left, right" & vbCrLf & vbCrLf & "player 1: UP arrow, DOWN arrow, LEFT arrow, RIGHT arrow" & vbCrLf & "player 2: W, S, A, D" & vbCrLf & "giocatore 3: keypad 4, 5, 2, 8" & vbCrLf & "player 4: I, K, J, L" & vbCrLf & vbCrLf & "F2: new game, F3: pause"
Global Const Version As String = "VBPython 1.08e"

Global MapValid() As Boolean
