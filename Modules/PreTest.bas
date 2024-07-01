Attribute VB_Name = "PreTest"
' Module for handling the logic and events within the PreTest.

' Declare variables for keeping track of score count..
Dim ScoreCorrect, ScoreIncorrect, ScoreGrade As Integer

' Initializes pretest elements.
Sub Initialize()
    Set Score = ActivePresentation.Slides(SlidePreResults)
    ' Sets pretest variables to 0.
    ScoreCorrect = 0
    ScoreIncorrect = 0
    ScoreGrade = 0
    CheckpointPretest = False
    
    ' Displays 0 on the results.
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = ScoreGrade & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPre").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPre").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePre").TextFrame.TextRange = ScoreGrade & "%"
    
    ' Calls subroutines to randomize answer order and shuffle slides.
    RandomizeAnswerOrder
    ShuffleSlides
End Sub

' Changes the color of the start button upon hovering.
Sub ResponseStartHover()
    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 217, 102)
End Sub

' Revers the start button color to white.
Sub ResponseStartHoverFalse()
    Set Slide = ActivePresentation.Slides(1)
    Slide.Shapes("ResponseStart").TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
End Sub

' Registers a correct answer to the score count.
Sub CorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePreResults)
    ' Adds a correct point to the score count.
    ScoreCorrect = (ScoreCorrect) + 1
    ' Displays the correct score on their respective elements. 
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPre").TextFrame.TextRange = ScoreCorrect
    ' Calls the OverallGrade subroutine to calculate score percentage.
    OverallGrade
End Sub

' Registers an incorrect answer to the score count.
Sub IncorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePreResults)
    ' Adds an incorrect point to the score count.
    ScoreIncorrect = (ScoreIncorrect) + 1
    ' Displays the incorrect score count on their respective elements. 
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPre").TextFrame.TextRange = ScoreIncorrect
    ' Calls the OverallGrade subroutine to calculate score percentage.
    OverallGrade
End Sub

' Calculates the score percentage.
Sub OverallGrade()
    Set Score = ActivePresentation.Slides(SlidePreResults) 
    ' Formula for the Total score variable.
    Total = ScoreCorrect + ScoreIncorrect
    ' Formula for computing the percentage.
    ScoreGrade = Round(ScoreCorrect / Total * 100, 1)
    ' Displays the score percentage on their respective elements.
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = (ScoreGrade) & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePre").TextFrame.TextRange = ScoreGrade & "%"
End Sub

' Shuffles slides.
Sub ShuffleSlides()
    FirstSlide = SlidePreFQ 
    LastSlide = SlidePreLQ  
    
    Randomize
    ' Generate random number for First and Last slides.
        RSN = Int((LastSlide - FirstSlide + 1) * Rnd + FirstSlide)
    
    ' Shuffles the slides.
    For i = FirstSlide To LastSlide
        ActivePresentation.Slides(i).MoveTo (RSN)
    Next i
End Sub

' Randomizes the answer order.
Sub RandomizeAnswerOrder()
    ' Declare variable to manage answer order.
    Dim AnswerOrder() As Integer
    ReDim AnswerOrder(3) ' Sets four compartments for the answer order. (Zero-indexed: 0, 1, 2, 3 --> 4 compartments)
    
    ' Initializes the randomizer.
    For i = 0 To 3
    AnswerOrder(i) = i + 1
    Next i
    
    ' Sets the slides to shuffle answers in.
    For i = SlidePreFQ To SlidePreLQ 
    
    ' Generates a random number.
    Randomize
    For n = 0 To 3
    j = Int(4 * Rnd) ' Random number from 0 to 3
    
    temp = AnswerOrder(n)
    AnswerOrder(n) = AnswerOrder(j)
    AnswerOrder(j) = temp
    
    Next n
    
    ' Moves the Choices to random positions with fixed coordinates.
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



