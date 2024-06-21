Attribute VB_Name = "PostTest"
Dim ScoreCorrect, ScoreIncorrect, ScoreGrade As Integer
Dim LastQuestion As Long


Sub Initialize()
    Set Score = ActivePresentation.Slides(SlidePostResults) 'Posttest results screen slide
    ScoreCorrect = 0
    ScoreIncorrect = 0
    ScoreGrade = 0
    
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = ScoreGrade & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPost").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPost").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePost").TextFrame.TextRange = ScoreGrade & "%"
    
    HPText = Round((15 - ScoreCorrect) / 15 * 100, 0)
    For i = SlidePostFQ To SlidePostLast 'Slides with boss hpbar
        For n = 1 To 15
            ActivePresentation.Slides(i).Shapes("!!HPBar" & n).Visible = True
            ActivePresentation.Slides(i).Shapes("!!HPBar" & n).Fill _
            .ForeColor.RGB = RGB(141, 193, 21)
        Next n
            ActivePresentation.Slides(i).Shapes("!!HPText").TextFrame.TextRange = HPText & "/100"
    Next i
    
    RandomizeAnswerOrder
    ShuffleSlides
End Sub

Sub CorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePostResults) 'Posttest results screen slide
    ScoreCorrect = (ScoreCorrect) + 1
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPost").TextFrame.TextRange = ScoreCorrect
    OverallGrade
    HPBarFunction
End Sub

Sub IncorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePostResults) 'Posttest results screen slide
    ScoreIncorrect = (ScoreIncorrect) + 1
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPost").TextFrame.TextRange = ScoreIncorrect
    OverallGrade
End Sub

Sub OverallGrade()
    Set Score = ActivePresentation.Slides(SlidePostResults) 'Posttest results screen slide
    Total = ScoreCorrect + ScoreIncorrect
    ScoreGrade = Round(ScoreCorrect / Total * 100, 1)
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = (ScoreGrade) & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePost").TextFrame.TextRange = ScoreGrade & "%"
End Sub

Sub HPBarFunction()
    Dim Damage As Integer
    Damage = ScoreCorrect
    HPText = Round((15 - ScoreCorrect) / 15 * 100, 0)
        On Error GoTo HandleError
        For i = SlidePostFQ To SlidePostLast 'Slides with boss hpbar
            ActivePresentation.Slides(i).Shapes("!!HPBar" & (Damage)).Visible = False
            ActivePresentation.Slides(i).Shapes("!!HPText").TextFrame.TextRange = HPText & "/100"
            If ActivePresentation.Slides(i).Shapes("!!HPBar8").Visible = False Then
                For n = 1 To 15
                    ActivePresentation.Slides(i).Shapes("!!HPBar" & (n)).Fill _
                    .ForeColor.RGB = RGB(229, 101, 46)
                Next n
            End If
            
            If ActivePresentation.Slides(i).Shapes("!!HPBar11").Visible = False Then
                For n = 1 To 15
                    ActivePresentation.Slides(i).Shapes("!!HPBar" & (n)).Fill _
                    .ForeColor.RGB = RGB(169, 46, 69)
                Next n
            End If
        Next i
        
HandleError:
    Resume Next
End Sub

Sub ShuffleSlides()
    FirstSlide = SlidePostFQ 'First question of pretest slide
    LastSlide = SlidePostLQ  'Last question of pretest slide
    
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
    
    For i = SlidePostFQ To SlidePostLQ 'Posttest question slides
    
    Randomize
    For n = 0 To 3
    j = Int(4 * Rnd) 'random number from 0 to 3
    
    temp = AnswerOrder(n)
    AnswerOrder(n) = AnswerOrder(j)
    AnswerOrder(j) = temp
    
    Next n
    
    For j = 0 To 3

        If AnswerOrder(j) = 1 Then
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Top = 372.5391
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Left = 757.6702
            
        ElseIf AnswerOrder(j) = 2 Then
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Top = 456.8076
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Left = 557.0136
            
        ElseIf AnswerOrder(j) = 3 Then
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Top = 372.6392
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Left = 558.3984
            
        ElseIf AnswerOrder(j) = 4 Then
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Top = 456.8076
            ActivePresentation.Slides(i).Shapes("!!Choice" & j + 1).Left = 757.3005
        End If

    Next j
    Next i
    Exit Sub
End Sub

