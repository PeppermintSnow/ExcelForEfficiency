Attribute VB_Name = "Checkpoints"
' Module for checkpoints; responsible for triggering and controlling certain events within the program.

' Declare global/public boolean variables for checking if a certain point is reached
Public CheckpointPrologueKeyResponse, CheckpointPretest, _
        CheckpointXenoluminaFV, CheckpointXenoluminaL1, CheckpointXenoluminaL2, CheckpointXenoluminaL3, CheckpointXenoluminaL4, CheckpointXenoluminaComplete, _
        CheckpointAuroraFV, CheckpointAuroraL1, CheckpointAuroraL2, CheckpointAuroraComplete, _
        CheckpointTenebrisAttack, _
        SampleCheck _
As Boolean

' Initializes the boolean variables.
' Checkpoint_FV checks if it is the first time the user visits the planet; FV = First Visit.
' Checkpoint_L# checks if the user has finished taking the lesson number; L1 = Lesson 1.
' Checkpoint_Complete checks if the user has finished all available lessons in the planet.
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

    ' Call PrologueProceed
    PrologueProceed
End Sub

' Checks if the user has taken the key dialogue in the intro.
Sub PrologueProceed()
    If CheckpointPrologueKeyResponse = True Then
        ActivePresentation.Slides(19).Shapes("!!Response4").Visible = msoTrue
    Else
        ActivePresentation.Slides(19).Shapes("!!Response4").Visible = msoFalse
    End If
End Sub

' Sets the dialogue key to true once the user has selected the key dialogue.
Sub PrologueKey()
    CheckpointPrologueKeyResponse = True
    ActivePresentation.SlideShowWindow.View.GotoSlide 24
    PrologueProceed
End Sub

' Sets the pretest checkpoint to true.
Sub ButtonPreTest()
    CheckpointPretest = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Handles the early-game dialogues in Tenebris.
Sub TenebrisAsk()
    If CheckpointXenoluminaL1 = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 257
    ElseIf CheckpointXenoluminaL1 = True And CheckpointXenoluminaComplete = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 258
    ElseIf CheckpointXenoluminaComplete = True Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 259
    End If
End Sub

' Handles the user's first visit in Xenolumina.
Sub XenoluminaFV()
    If CheckpointXenoluminaFV = True Then
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideXenoluminaFV
    ElseIf CheckpointXenoluminaFV = False Then
        ActivePresentation.SlideShowWindow.View.GotoSlide (SlideXenoluminaMenu + 1)
    End If
End Sub

' Sets Xenolumina's first visit checkpoint to false.
Sub ButtonXenoluminaFV() 'Slide 93
    CheckpointXenoluminaFV = False
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Xenolumina's Lesson 1 checkpoint to true.
Sub ButtonXenoluminaL1() 'Slide 107
    CheckpointXenoluminaL1 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Xenolumina's Lesson 2 checkpoint to true.
Sub ButtonXenoluminaL2() 'Slide 122
    CheckpointXenoluminaL2 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Xenolumina's Lesson 3 checkpoint to true.
Sub ButtonXenoluminaL3() 'Slide 138
    CheckpointXenoluminaL3 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Xenolumina's Lesson 4 checkpoint to true.
Sub ButtonXenoluminaL4() 'Slide 149
    CheckpointXenoluminaL4 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Handles the user's first visit in Aurora.
Sub ButtonAuroraFV()
    CheckpointAuroraFV = False
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Aurora's Lesson 1 checkpoint to true.
Sub ButtonAuroraL1() 'Slide 215
    CheckpointAuroraL1 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Sets Aurora's Lesson 1 checkpoint to true.
Sub ButtonAuroraL2() 'Slide 227
    CheckpointAuroraL2 = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Subroutines for quick fixes and debugging in case of errors.
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
