Attribute VB_Name = "FindObjectPosition"
Sub FindPos()
    i = 35
    Debug.Print ActivePresentation.Slides(i).Shapes("DialogueBox").Top
    Debug.Print ActivePresentation.Slides(i).Shapes("DialogueBox").Left
End Sub

Sub FindSize()
    i = 74
    For n = 1 To 4
    Debug.Print ActivePresentation.Slides(i).Shapes("!!Choice" & n).Height
    Debug.Print ActivePresentation.Slides(i).Shapes("!!Choice" & n).Width
    Next n
End Sub

