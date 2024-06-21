Attribute VB_Name = "PreTest"
Dim ScoreCorrect, ScoreIncorrect, ScoreGrade As Integer

Sub Initialize()
    Set Score = ActivePresentation.Slides(SlidePreResults)
    ScoreCorrect = 0
    ScoreIncorrect = 0
    ScoreGrade = 0
    CheckpointPretest = False
    
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = ScoreGrade & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPre").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPre").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePre").TextFrame.TextRange = ScoreGrade & "%"
    
    RandomizeAnswerOrder
    ShuffleSlides
End Sub

Sub ResponseStartHover()
    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
End Sub

Sub ResponseStartHoverFalse()
    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
End Sub

Sub CorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePreResults) 'Results screen slide
    ScoreCorrect = (ScoreCorrect) + 1
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPre").TextFrame.TextRange = ScoreCorrect
    OverallGrade
End Sub

Sub IncorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePreResults) 'Results screen slide
    ScoreIncorrect = (ScoreIncorrect) + 1
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPre").TextFrame.TextRange = ScoreIncorrect
    OverallGrade
End Sub

Sub OverallGrade()
    Set Score = ActivePresentation.Slides(SlidePreResults) 'Results screen slide
    Total = ScoreCorrect + ScoreIncorrect
    ScoreGrade = Round(ScoreCorrect / Total * 100, 1)
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = (ScoreGrade) & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePre").TextFrame.TextRange = ScoreGrade & "%"
End Sub

Sub ShuffleSlides()
    FirstSlide = SlidePreFQ 'First question of pretest slide
    LastSlide = SlidePreLQ  'Last question of pretest slide
    
    Randomize
    'generate random number between 2 to 7'
        RSN = Int((LastSlide - FirstSlide + 1) * Rnd + FirstSlide)
    
    For i = FirstSlide To LastSlide
        ActivePresentation.Slides(i).MoveTo (RSN)
    Next i
End Sub

Sub RandomizeAnswerOrder()
    Dim AnswerOrder() As Integer
    ReDim AnswerOrder(3) '0 1 2 3 -> 4 compartments'
    
    For i = 0 To 3
    AnswerOrder(i) = i + 1
    Next i
    
    For i = SlidePreFQ To SlidePreLQ 'Pretest questions slides
    
    Randomize
    For n = 0 To 3
    j = Int(4 * Rnd) 'random number from 0 to 3
    
    temp = AnswerOrder(n)
    AnswerOrder(n) = AnswerOrder(j)
    AnswerOrder(j) = temp
    
    Next n
    
    For j = 0 To 3
    
    On Error GoTo HandleError
        If AnswerOrder(j) = 1 Then
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Top = 353.2105
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Left = 86.15606
            
        ElseIf AnswerOrder(j) = 2 Then
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Top = 353.1104
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Left = 296.5051
            
        ElseIf AnswerOrder(j) = 3 Then
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Top = 437.3789
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Left = 86.15606
            
        ElseIf AnswerOrder(j) = 4 Then
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Top = 437.3789
            ActivePresentation.Slides(i).Shapes("Choice" & j + 1).Left = 296.5051
        End If

    Next j
    Next i
    Exit Sub
    
HandleError:
    Resume Next
    
End Sub

Sub test()
Debug.Print ScoreIncorrect
End Sub



