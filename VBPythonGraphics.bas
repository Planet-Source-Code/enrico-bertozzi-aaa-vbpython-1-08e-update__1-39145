Attribute VB_Name = "VBPythonGraphics"
Option Explicit

'Private Declare Function Playsound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long 'filename,0,1
'NO SOUNDS YET!
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCAND = &H8800C6

'Public Sub WPlaySound(FileName As String)

'Dim S&
'S = Playsound(FileName, 0, 1)

'End Sub

Public Function DrwTranspSpriteBlt(pTo As Object, pToX As Integer, pToY As Integer, pFrom As Object, pMask As Object)

Static E&
E = BitBlt(pTo.hDC, pToX, pToY, pFrom.hDC, pFrom.Height, pMask.hDC, 0, 0, SRCAND)
DrwTranspSpriteBlt = E And BitBlt(pTo.hDC, pToX, pToY, pFrom.Width, pFrom.Height, pFrom.hDC, 0, 0, SRCINVERT)

End Function

'Public Function DrwSpriteBlt(pToWnd As Long, pToX As Integer, pToY As Integer, pFromWnd As Long)

'Static E&
'E = BitBlt(pToWnd, pToX, pToY, pFrom.Width, pFrom.Height, pFromWnd, 0, 0, SRCCOPY)

'End Function

Sub Main()
Dim lowResAns As Integer

If Screen.Width \ Screen.TwipsPerPixelX < 1023 Then
    lowResAns = MsgBox("This game plays best at 1024x768 or higher resolution. Continue anyway?", vbYesNo + vbExclamation)
    If lowResAns = vbNo Then End
End If

frmMain.Show
End Sub
