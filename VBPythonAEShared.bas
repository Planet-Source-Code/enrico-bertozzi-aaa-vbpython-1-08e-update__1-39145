Attribute VB_Name = "VBPythonAEShared"
Sub FldClear()

'here we set all the blocks to -2, i.e. to a blank tile
Dim I%, J%
For I = 1 To 88
    For J = 1 To 63
        FldX(I, J) = -2
        FldY(I, J) = -2
    Next J
Next I
End Sub

Sub MapDraw()
Dim I%, J%

'simple: we paint each block with its wall part. Specifying a negative value for
'[Height1], [Width1] etc., means that the sprite will be flipped (X or Y)

'PaintPicture method is slow, especially with complex maps
For I = 1 To 88
    For J = 1 To 63
        Select Case FldX(I, J)
            Case fldMCross, fldMV, fldMH
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres" & FldX(I, J) & ".bmp"), QX(I), QY(J)
            Case fldMI1
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-7.bmp"), QX(I), QY(J)
            Case fldMI2
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-9.bmp"), QX(I), QY(J)
            Case fldMI3
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-7.bmp"), QX(I), QY(J) + 9, 10, -10
            Case fldMI4
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-9.bmp"), QX(I) + 9, QY(J), -10, 10
            Case fldMRot1
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), QX(I), QY(J)
            Case fldMRot2
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), QX(I) + 9, QY(J), -10, 10
            Case fldMRot3
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), QX(I) + 9, QY(J) + 9, -10, -10
            Case fldMRot4
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\wres-3.bmp"), QX(I), QY(J) + 9, 10, -10
            Case -1, 0, 1
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\py1s.bmp"), QX(I), QY(J)
            Case 2, 3, 4
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\py2s.bmp"), QX(I), QY(J)
            Case 5, 6, 7
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\py3s.bmp"), QX(I), QY(J)
            Case 8, 9, 10
            frmMain.pGame.PaintPicture LoadPicture(App.Path + "\py4s.bmp"), QX(I), QY(J)
        End Select
    Next J
Next I
End Sub

Function MapSet(FileName As String) As Boolean
Dim ff As Integer, tmp As String * 1, tmp2%, I%, J%

On Error GoTo ErrHnd

ff = FreeFile
J = 1
I = 1
'read the mapfile, and convert all ASCII to field values < -2
Open App.Path + "\" + FileName For Input As #ff
Do
    tmp = Input(ff, 1)
    tmp2 = (Asc(tmp) - 63) * -1
    If tmp2 > -2 Or tmp2 < -13 Then
        MapSet = False
        Close
        Exit Function
    End If
    FldX(I, J) = tmp2
    I = I + 1
    If I = 89 Then I = 1: J = J + 1
    Loop Until EOF(ff) Or Seek(ff) > 5544
Close #ff

MapSet = True
Exit Function

ErrHnd:
MapSet = False
End Function

Sub QSet()

For I = 1 To 88
    QX(I) = (I - 1) * 10
Next I
For I = 1 To 63
    QY(I) = (I - 1) * 10
Next I

End Sub

Sub DrwSpriteBlt(pTo As Object, pToX As Integer, pToY As Integer, pFrom As Object)

BitBlt pTo.hdc, pToX, pToY, pFrom.Width, pFrom.Height, pFrom.hdc, 0, 0, vbSrcCopy

End Sub

Function DetectMaps() As Integer
Dim MDetect$, IMDetect%, I%

'previous version (1.0) would accept only N maps, 1 to N, and
'had issues with corrupted (shorter and longer) maps.

'this, instead, supports any map number, even not sorted, and
'detects corrupt mapfiles, avoiding selecting them.

MDetect = Dir(App.Path & "\map*.pm")
Do While Not MDetect = "" 'determine highest map number
    MDetect = Left(MDetect, Len(MDetect) - 3)
    MDetect = Right(MDetect, Len(MDetect) - 3)
    If CInt(MDetect) > IMDetect Then IMDetect = CInt(MDetect)
    MDetect = Dir
Loop

'mapvalid is a dynamic-matrix, used to determine if a map is valid or not.
'We set a specific cell to TRUE if the map exists and is valid. Else, set it
'FALSE
ReDim MapValid(1 To IMDetect)

For I = 1 To IMDetect
    On Error GoTo ErrHnd
    If FileLen(App.Path & "\map" & I & ".pm") < 5544 Then MapValid(I) = False Else MapValid(I) = True
    'if FileLen is less than 5544, the map is corrupted and will not be used
Next I

DetectMaps = IMDetect
Exit Function

ErrHnd:
If Err.Number = 53 Then
    'if in the FileLen intruction there is an error (53=file not found), then
    'the map doesn't exist
    MapValid(I) = False
    'resume the next instruction after the error
    Resume Next
End If
End Function
