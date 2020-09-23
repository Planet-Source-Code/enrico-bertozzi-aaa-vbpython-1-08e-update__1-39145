Attribute VB_Name = "AEMisc"
Sub SaveMap(newNumber As Integer)
Dim I%, J%, ff%, temp As Byte

ff = FreeFile
Open App.Path & "\map" & newNumber & ".pm" For Output As #ff

For I = 1 To 63
    For J = 1 To 88
        temp = FldX(J, I) * -1 + 63
        Print #ff, Chr(temp);
    Next J
Next I

Close #ff
End Sub

Function GetSaveLoc() As Integer
Dim I%

For I = 1 To IMDetect
    If Not MapValid(I) Then
        GetSaveLoc = I
        Exit Function
    End If
Next I

GetSaveLoc = IMDetect + 1
End Function

Sub AddSnakes()
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
End Sub

Sub RemoveSnakes()
    FldX(9, 5) = -2
    FldX(8, 5) = -2
    FldX(7, 5) = -2
    FldX(6, 5) = -2
    FldX(5, 5) = -2
    FldY(5, 5) = -2
    FldY(6, 5) = -2
    FldY(7, 5) = -2
    FldY(8, 5) = -2
    FldY(9, 5) = -2
    FldX(80, 5) = -2
    FldX(81, 5) = -2
    FldX(82, 5) = -2
    FldX(83, 5) = -2
    FldX(84, 5) = -2
    FldY(80, 5) = -2
    FldY(81, 5) = -2
    FldY(82, 5) = -2
    FldY(83, 5) = -2
    FldY(84, 5) = -2
    FldX(9, 59) = -2
    FldX(8, 59) = -2
    FldX(7, 59) = -2
    FldX(6, 59) = -2
    FldX(5, 59) = -2
    FldY(9, 59) = -2
    FldY(8, 59) = -2
    FldY(7, 59) = -2
    FldY(6, 59) = -2
    FldY(5, 59) = -2
    FldX(80, 59) = -2
    FldX(81, 59) = -2
    FldX(82, 59) = -2
    FldX(83, 59) = -2
    FldX(84, 59) = -2
    FldY(80, 59) = -2
    FldY(81, 59) = -2
    FldY(82, 59) = -2
    FldY(83, 59) = -2
    FldY(84, 59) = -2
End Sub
