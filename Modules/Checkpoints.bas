Attribute VB_Name = "Checkpoints"
'Variables for checking if a certain point is reached
Public CheckpointPrologueKeyResponse, CheckpointPretest, _
        CheckpointXenoluminaFV, CheckpointXenoluminaL1, CheckpointXenoluminaL2, CheckpointXenoluminaL3, CheckpointXenoluminaL4, CheckpointXenoluminaComplete, _
        CheckpointAuroraFV, CheckpointAuroraL1, CheckpointAuroraL2, CheckpointAuroraComplete, _
        CheckpointTenebrisAttack, _
        SampleCheck _
As Boolean

Sub InitializeCheckpoints()
    CheckpointPrologueKeyResponse = False
    CheckpointPretest = False
    CheckpointXenoluminaFV = True
    CheckpointXenoluminaL1 = False
    CheckpointXenoluminaL2 = False
    CheckpointXenoluminaL3 = False
    CheckpointXenoluminaL4 = False
    CheckpointXenoluminaComplete = False
    CheckpointAuroraFV = True
    CheckpointAuroraL1 = False
    CheckpointAuroraL2 = False
    CheckpointAuroraComplete = False
    CheckpointTenebrisAttack = False
    PrologueProceed
End Sub

Sub PrologueProceed()
    If CheckpointPrologueKeyResponse = True Then
        ActivePresentation.Slides(19).Shapes("!!Response4").Visible = msoTrue
    Else
        ActivePresentation.Slides(19).Shapes("!!Response4").Visible = msoFalse
    End If
End Sub
Sub PrologueKey()
    CheckpointPrologueKeyResponse = True
    ActivePresentation.SlideShowWindow.View.GotoSlide 24
    PrologueProceed
End Sub

Sub ButtonPreTest()
    CheckpointPretest = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub TenebrisAsk()
    If CheckpointXenoluminaL1 = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 257
    ElseIf CheckpointXenoluminaL1 = True And CheckpointXenoluminaComplete = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 258
    ElseIf CheckpointXenoluminaComplete = True Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 259
    End If
End Sub

Sub XenoluminaFV()
    If CheckpointXenoluminaFV = True Then
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideXenoluminaFV
    ElseIf CheckpointXenoluminaFV = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (SlideXenoluminaMenu + 1)
    End If
End Sub

Sub ButtonXenoluminaFV() 'Slide 93
    CheckpointXenoluminaFV = False
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonXenoluminaL1() 'Slide 107
    CheckpointXenoluminaL1 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonXenoluminaL2() 'Slide 122
    CheckpointXenoluminaL2 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonXenoluminaL3() 'Slide 138
    CheckpointXenoluminaL3 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonXenoluminaL4() 'Slide 149
    CheckpointXenoluminaL4 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonAuroraFV()
    CheckpointAuroraFV = False
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonAuroraL1() 'Slide 215
    CheckpointAuroraL1 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub ButtonAuroraL2() 'Slide 227
    CheckpointAuroraL2 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub test()
    Debug.Print CheckpointXenoluminaL1
    Debug.Print CheckpointXenoluminaL2
    Debug.Print CheckpointXenoluminaL3
    Debug.Print CheckpointXenoluminaL4
End Sub

Private Sub CodeQuickFix()
    CheckpointPretest = True
    CheckpointXenoluminaComplete = True
    CheckpointAuroraFV = True
    Debug.Print CheckpointXenoluminaComplete
End Sub

Private Sub CodeDebug()
    Debug.Print CheckpointXenoluminaComplete
    Debug.Print CheckpointAuroraFV
End Sub
