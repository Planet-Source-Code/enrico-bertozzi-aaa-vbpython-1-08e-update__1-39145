Attribute VB_Name = "VBPythonIO"
Sub ReadHS()
Dim I%, ff%

On Error GoTo ExitSub
ff = FreeFile
Open App.Path & "\scores.ini" For Input As #ff
For I = 1 To 5
    Line Input #ff, hSNames(I)
    Line Input #ff, hSPoints(I)
Next I
ExitSub:
Close #ff
End Sub

Sub WriteHS()
Dim I%, ff%

ff = FreeFile
Open App.Path & "\scores.ini" For Output As #ff
For I = 1 To 5
    Print #ff, hSNames(I)
    Print #ff, hSPoints(I)
Next I
Close #ff
End Sub

