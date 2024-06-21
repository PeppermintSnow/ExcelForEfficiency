Attribute VB_Name = "PageDetector"
'Variables for posttest question return after animation
Dim LastQuestion As Long
Dim Question As Integer

'Variable for detecting current slide
Public CurrentSlide As Integer

'Variables for collective gathering of major slides
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


Sub OnSlideShowPageChange()
    CurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    
    If CurrentSlide = 3 Then 'Menu screen slide
        PreTest.Initialize
        CheckpointPretest = False
        CheckpointOFIntro = False
        CheckpointOFOSample = False
        
    ElseIf CurrentSlide = 84 Then 'Assessment loading screen slide
        PostTest.Initialize
        Question = 0

    ElseIf CurrentSlide = 74 Then
        If CheckpointXenoluminaFV = False Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 158
        Else
        End If
    
    ElseIf CurrentSlide = 157 Then 'Transtion
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideXenoluminaMenu
    
    ElseIf CurrentSlide = 109 Or CurrentSlide = 124 Or CurrentSlide = 140 Or CurrentSlide = 151 Or CurrentSlide = 154 Then 'Xenolumina completion dialogue
        If CheckpointXenoluminaComplete = False Then
            If CheckpointXenoluminaL1 = True And CheckpointXenoluminaL2 = True And CheckpointXenoluminaL3 = True And CheckpointXenoluminaL4 = True Then
                CheckpointXenoluminaComplete = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 155
            Else
                CheckpointXenoluminaComplete = False
            End If
        Else
        End If
        
    ElseIf CurrentSlide = 172 Then 'Aurora requirements checker
        If CheckpointAuroraFV = True Then
            If CheckpointXenoluminaComplete = True Then
                ActivePresentation.SlideShowWindow.View.GotoSlide 173
            Else
                ActivePresentation.SlideShowWindow.View.GotoSlide 188
            End If
        Else
            If CheckpointXenoluminaComplete = True Then
                ActivePresentation.SlideShowWindow.View.GotoSlide 173
            Else
                ActivePresentation.SlideShowWindow.View.GotoSlide 188
            End If
        End If
        
    ElseIf CurrentSlide = 173 Then
        If CheckpointAuroraFV = False Then
            If CheckpointXenoluminaComplete = True Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 234
            Else
            End If
        Else
        End If
        
    ElseIf CurrentSlide = 185 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 195
        
    ElseIf CurrentSlide = 196 Or CurrentSlide = 219 Or CurrentSlide = 231 Then
        If CheckpointAuroraComplete = False Then
            If CheckpointAuroraL1 = True And CheckpointAuroraL2 = True Then
                CheckpointAuroraComplete = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 197
            Else
                CheckpointAuroraComplete = False
            End If
        End If
    
    ElseIf CurrentSlide = 201 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 194
    
    ElseIf CurrentSlide = 242 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 244
    ElseIf CurrentSlide = 72 Or CurrentSlide = 170 Or CurrentSlide = 245 Then
        If CheckpointTenebrisAttack = False Then
            If CheckpointAuroraComplete = True Then
                CheckpointTenebrisAttack = True
                ActivePresentation.SlideShowWindow.View.GotoSlide 246
            Else
                CheckpointTenebrisAttack = False
            End If
        End If
        
    ElseIf CurrentSlide = 248 Then
        If CheckpointTenebrisAttack = True Then
            ActivePresentation.Slides(248).Shapes("!!LabelWarning").Visible = msoTrue
        Else
            ActivePresentation.Slides(248).Shapes("!!LabelWarning").Visible = msoFalse
        End If
            
    ElseIf CurrentSlide = 249 Then
        If CheckpointTenebrisAttack = True Then
            ActivePresentation.SlideShowWindow.View.GotoSlide 260
        Else
            ActivePresentation.SlideShowWindow.View.GotoSlide 250
        End If
        
    ElseIf CurrentSlide = 281 Then
        PostTest.CorrectAnswer
        
    ElseIf CurrentSlide = 277 Or CurrentSlide = 285 Then
        PostTestFinish
        
    ElseIf CurrentSlide = 284 Or CurrentSlide = 290 Then
        ReturnToNextQuestion
        Question = (Question + 1)
        PostTestFinish
    ElseIf CurrentSlide = 291 Or CurrentSlide = 293 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 294
    
    ElseIf CurrentSlide = 317 Then
        ResultAssessment
    End If
        
    
End Sub

Sub ResultAssessment()
    Dim INCorDEC As String
    Dim Percentage
    Set PreGrade = ActivePresentation.Slides(SlidePreResults).Shapes("!!VBoxGrade").TextFrame.TextRange
    Set PostGrade = ActivePresentation.Slides(SlidePostResults).Shapes("!!VBoxGrade").TextFrame.TextRange
    
    Percentage = PostGrade - PreGrade
    
    If PostGrade > PreGrade Then
        INCorDEC = "an increase"
    Else
        INCorDEC = "a decrease"
    End If
    
    Set Slide = ActivePresentation.Slides(318)
    Slide.Shapes("!!BoxInterpretation").TextFrame.TextRange = "By comparing your pre-assessment and post-assessment scores, " & INCorDEC & " by " & Percentage & "% has been observed in your performance. Thank you for using Excel For Efficiency!"
End Sub
Sub InitializeAll()

    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    
    Checkpoints.InitializeCheckpoints
    PreTest.Initialize
    PostTest.Initialize
    ResetResponseColor
    ActivePresentation.SlideShowWindow.View.Next

End Sub

Sub ResetResponseColor()
    On Error GoTo HandleError:
        For i = 1 To 318
            For n = 1 To 5
                ActivePresentation.Slides(i).Shapes("!!Response" & n).TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            Next n
        Next i
    Exit Sub
HandleError:
    Resume Next
End Sub
            
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

Sub ResponseHover4()
On Error GoTo HandleError
    Set Slide = ActivePresentation.Slides(CurrentSlide)
    Slide.Shapes("!!Response1").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response2").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response3").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Slide.Shapes("!!Response4").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
    Slide.Shapes("!!Response5").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
End Sub

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

Sub CorrectRememberLastQuestion()
    LastQuestion = ActivePresentation.SlideShowWindow.View.Slide.slideIndex
    ActivePresentation.SlideShowWindow.View.GotoSlide 277 'Bob first attack frame slide
End Sub

Sub IncorrectRememberLastQuestion()
    LastQuestion = ActivePresentation.SlideShowWindow.View.Slide.slideIndex
    ActivePresentation.SlideShowWindow.View.GotoSlide 285 'Bob first attack frame slide
    PostTest.IncorrectAnswer
End Sub

Sub ReturnToNextQuestion()
    If Question < 14 Then
        With SlideShowWindows(1).View
        .GotoSlide (LastQuestion + 1)
        End With
    End If
End Sub

Sub PostTestFinish()
    If ActivePresentation.Slides(297).Shapes("!!BoxCorrect").TextFrame.TextRange = "15" And Question > 14 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 292
    ElseIf Question > 14 Then
        ActivePresentation.SlideShowWindow.View.GotoSlide 293
    End If
End Sub

Sub test()
    Debug.Print Question
End Sub
