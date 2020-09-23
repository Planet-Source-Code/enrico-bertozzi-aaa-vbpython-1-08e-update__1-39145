Attribute VB_Name = "AEDims"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Global IMDetect As Integer, CurrentMap As Integer

Global QX(1 To 88) As Integer
Global QY(1 To 63) As Integer
Global FldX(1 To 88, 1 To 63) As Integer
Global FldY(1 To 88, 1 To 63) As Integer

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

Global Selected As Integer

Global MapValid() As Boolean

Global Const Version As String = "Ambient Editor beta 1.01d"

Global FileSave As Boolean, FileNew As Boolean
