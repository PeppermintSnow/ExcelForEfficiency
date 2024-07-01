Attribute VB_Name = "PageDetector"
' Module responsible for detecting the pages and handling other events.

' Declare variables for the posttest attack animation.
Dim LastQuestion As Long
Dim Question As Integer

' Declare a public variable for detectin the current slide.
Public CurrentSlide As Integer

' Declare public variables to easily adjust the slide number of important slides
Public Const SlideXenoluminaTransition As Integer = 74
Public Const SlidePreResults As Integer = 52
    Public Const SlidePreFQ As Integer = 35
    Public Const SlidePreLQ As Integer = 49
Public Const SlideXenoluminaLessonMenu As Integer = 154
    Public Const SlideOFTransition As Integer = 33
    Public Const SlideOFBLast As Integer = 47
    Public Const SlideOFOLast As Integer = 62
    Public Const SlideOFFLast As Integer = 78
    Public Const SlideOFIFLast As Integer = 89
Public Const SlidePostResults As Integer = 296
    Public Const SlidePostFQ As Integer = 262
    Public Const SlidePostLQ As Integer = 276
    Public Const SlidePostLast As Integer = 294
Public Const SlideFinalResults As Integer = 318
Public Const SlideXenoluminaFV = 75
Public Const SlideXenoluminaMenu = 159
Public Const SlideAuroraEV = 188
Public Const SlideAuroraFV = 173
Public Const SlideAuroraMenu = 230
Public Const SlideTenebrisEV = 247
Public Const SlideTenebrisWarning = 245
Public Const SlideTenebrisBattle = 261

' Detects slide number.
Sub OnSlideShowPageChange()
    ' Sets the CurrentSlide variable as the number of the current slide.
    CurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    
    ' Initializes the pretest and Xenolumina.
    If CurrentSlide = 3 Then 'Menu screen slide
        ' Calls the Initialize subroutine from the PreTest module.
        PreTest.Initialize
        ' Initializes the checkpoints for pretest and Xenolumina
        CheckpointPretest = False
        CheckpointOFIntro = False
        CheckpointOFOSample = False
        
    ' Initializes the posttest.
    ElseIf CurrentSlide = 84 Then 'Assessment loading screen slide
        ' Calls the Initialize subroutine from the PostTest module.
        PostTest.Initialize
        ' Initializes the Question variable for the posttest attack animation.
        Question = 0

    ' Runs a conditional statement to check if Xenolumina's first visit is true.
    ElseIf CurrentSlide = 74 Then
        If CheckpointXenoluminaFV = False Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 158
        Else
        End If
    
    ' Redirects the next slide for a smooth transition.
    ElseIf CurrentSlide = 157 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideXenoluminaMenu
    
    ' Checks if all lessons in Xenolumina is completed everytime the user visits the menu.c
    ElseIf CurrentSlide = 109 Or CurrentSlide = 124 Or CurrentSlide = 140 Or CurrentSlide = 151 Or CurrentSlide = 154 Then 
        ' Runs another conditional statement if the Complete variable is false.
        If CheckpointXenoluminaComplete = False Then
            ' Assigns the Complete variable to true if all four lessons are completed.
            If CheckpointXenoluminaL1 = True And CheckpointXenoluminaL2 = True And CheckpointXenoluminaL3 = True And CheckpointXenoluminaL4 = True Then
                CheckpointXenoluminaComplete = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 155
            ' Else it does nothing.
            Else
                CheckpointXenoluminaComplete = False
            End If
        Else
        End If
        
    ' Plays Aurora's dialogue if all prerequisite requirements are met. 
    ElseIf CurrentSlide = 172 Then 
        ' If First Visit is true, run another conditional statement.
        If CheckpointAuroraFV = True Then
            If CheckpointXenoluminaComplete = True Then
                ActivePresentation.SlideShowWindow.View.GotoSlide 173
            Else
                ActivePresentation.SlideShowWindow.View.GotoSlide 188
            End If
        ' Else, check if Xenolumina's Complete variable is true, then run another conditional statement.
        Else
            If CheckpointXenoluminaComplete = True Then
                ActivePresentation.SlideShowWindow.View.GotoSlide 173
            Else
                ActivePresentation.SlideShowWindow.View.GotoSlide 188
            End If
        End If
        
    ' If Aurora's First Visit is false, check if Xenolumina's Complete variable is true, then redirect the user.
    ElseIf CurrentSlide = 173 Then
        If CheckpointAuroraFV = False Then
            If CheckpointXenoluminaComplete = True Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 234
            Else
            End If
        Else
        End If
        
    ' Redirects the next slide for a smooth transition.
    ElseIf CurrentSlide = 185 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 195
        
    ' Checks if all lessons in Aurora is completed everytime the user visits the menu.
    ElseIf CurrentSlide = 196 Or CurrentSlide = 219 Or CurrentSlide = 231 Then
        ' If Aurora's Complete variable is false, check if all lessons are set to true; then set to true.
        If CheckpointAuroraComplete = False Then
            If CheckpointAuroraL1 = True And CheckpointAuroraL2 = True Then
                CheckpointAuroraComplete = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 197
            ' Otherwise, do nothing.
            Else
                CheckpointAuroraComplete = False
            End If
        End If
    
    ' Redirects the next slide for a smooth transition.
    ElseIf CurrentSlide = 201 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 194
    
    ' Redirects the next slide for a smooth transition.
    ElseIf CurrentSlide = 242 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 244

    ' Checks if Aurora is completed everytime the user visits the planets menu.
    ElseIf CurrentSlide = 72 Or CurrentSlide = 170 Or CurrentSlide = 245 Then
        ' Set Tenebris' Attack checkpoint to true if Aurora is complete.
        If CheckpointTenebrisAttack = False Then
            If CheckpointAuroraComplete = True Then
                CheckpointTenebrisAttack = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 246
            ' Otherwise, do nothing.
            Else
                CheckpointTenebrisAttack = False
            End If
        End If
    
    ' Sets the warning label to true if Tenebris' Attack variable is true.
    ElseIf CurrentSlide = 248 Then
        If CheckpointTenebrisAttack = True Then
            ActivePresentation.Slides(248).Shapes("!!LabelWarning").Visible = msoTrue
        Else
            ActivePresentation.Slides(248).Shapes("!!LabelWarning").Visible = msoFalse
        End If
            
    ' Redirects the user to the appropriate slide corresponding to their current progress.
    ElseIf CurrentSlide = 249 Then
        If CheckpointTenebrisAttack = True Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 260
        Else
            ActivePresentation.SlideShowWindow.View.GotoSlide 250
        End If
        
    ' Calls the CorrectAnswer subroutine from the PostTest module when reaching a certain slide.
    ElseIf CurrentSlide = 281 Then
        PostTest.CorrectAnswer
        
    ' Calls the PostTestFinish subroutine when reaching a certain slide.
    ElseIf CurrentSlide = 277 Or CurrentSlide = 285 Then
        PostTestFinish
        
    ' Responsible for running the posttest properly.
    ElseIf CurrentSlide = 284 Or CurrentSlide = 290 Then
        ' Once the animation is finished playing, return to the next question.
        ReturnToNextQuestion
        ' Sets the question variable to correctly redirect the slide number.
        Question = (Question + 1)
        ' Calls the PostTestFinish subroutine to check if done.
        PostTestFinish

    ' Redireccts the next slide for a smooth transition.
    ElseIf CurrentSlide = 291 Or CurrentSlide = 293 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 294
    
    ' Calls the ResultAssessment subroutine to calculate the results.
    ElseIf CurrentSlide = 317 Then
        ResultAssessment
    End If
        
    
End Sub

' Calculates results.
Sub ResultAssessment()
    ' Declare variables for displaying results.
    Dim INCorDEC As String
    Dim Percentage
    Set PreGrade = ActivePresentation.Slides(SlidePreResults).Shapes("!!VBoxGrade").TextFrame.TextRange
    Set PostGrade = ActivePresentation.Slides(SlidePostResults).Shapes("!!VBoxGrade").TextFrame.TextRange
    
    ' Formula for computing the percentage.
    Percentage = PostGrade - PreGrade
    
    ' Conditional statement to display the appropriate string based on the conditions.
    If PostGrade > PreGrade Then
        INCorDEC = "an increase"
    Else
        INCorDEC = "a decrease"
    End If
    
    ' Displays the appropriate string for interpretation of results.
    Set Slide = ActivePresentation.Slides(318)
    Slide.Shapes("!!BoxInterpretation").TextFrame.TextRange = "By comparing your pre-assessment and post-assessment scores, " & INCorDEC & " by " & Percentage & "% has been observed in your performance. Thank you for using Excel For Efficiency!"
End Sub

' Initializes all variables.
Sub InitializeAll()
    ' Displays the correct color for the start button in the first slide.
    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    
    ' Calls several initializing subroutines from different modules.
    Checkpoints.InitializeCheckpoints
    PreTest.Initialize
    PostTest.Initialize
    ResetResponseColor
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Reverts the color of the response buttons to white.
Sub ResetResponseColor()
    ' In case the shape is not detected, handle the exception by jumping to HandleError.
    On Error GoTo HandleError:
        ' Loop to run in slides 1 to 318.
        For i = 1 To 318
            For n = 1 To 5
                ActivePresentation.Slides(i).Shapes("!!Response" & n).TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Next n
        Next i
    Exit Sub
' Handles the errors thrown by resuming next.
HandleError:
    Resume Next
End Sub
        
' Changes the color of Response1 to yellow when hovered over.
Sub ResponseHover1()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Exit Sub
HandleError:
    Resume Next
End Sub

' Changes the color of Response2 to yellow when hovered over.
Sub ResponseHover2()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Exit Sub
HandleError:
    Resume Next
End Sub

' Changes the color of Response3 to yellow when hovered over.
Sub ResponseHover3()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Exit Sub
HandleError:
    Resume Next
End Sub

' Changes the color of Response4 to yellow when hovered over.
Sub ResponseHover4()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
End Sub

' Changes the color of Response5 to yellow when hovered over.
Sub ResponseHover5()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Exit Sub
HandleError:
    Resume Next
End Sub

' Resets the color of Response buttons to white.
Sub ResponseHoverFalse()
    On Error GoTo HandleError
        Set Slide = ActivePresentation.Slides(CurrentSlide)
        Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Exit Sub
HandleError:
    Resume Next
End Sub

' Stores the correct index for the LastQuestion variable, then redirects to the player attack animation. 
Sub CorrectRememberLastQuestion()
    LastQuestion = ActivePresentation.SlideShowWindow.View.Slide.slideIndex
    ActivePresentation.SlideShowWindow.View.GotoSlide 277 'Bob's first attack frame slide
End Sub

' Stores the correct index for the LastQuestion variable, then redirects to the boss attack animation. 
Sub IncorrectRememberLastQuestion()
    LastQuestion = ActivePresentation.SlideShowWindow.View.Slide.slideIndex
    ActivePresentation.SlideShowWindow.View.GotoSlide 285 'Boss' first attack frame slide
    ' Calls the IncorrectAnswer subroutine from the PostTest module to register an incorrect score.
    PostTest.IncorrectAnswer
End Sub

' Returns to the next question after playing the attack animation.
Sub ReturnToNextQuestion()
    ' If the Question index is less than 14, redirect to the next question.
    If Question < 14 Then
        With SlideShowWindows(1).View
        .GotoSlide (LastQuestion + 1)
        End With
    End If
End Sub

' Checks if the posttest is completed.
Sub PostTestFinish()
    ' If the user's score is perfect, run the boss' death animation.
    If ActivePresentation.Slides(297).Shapes("!!BoxCorrect").TextFrame.TextRange = "15" And Question > 14 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 292
    ' If the Question index is more than 14, redirect to the results screen.
    ElseIf Question > 14 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 293
    End If
End Sub

Sub test()
    Debug.Print Question
End Sub
