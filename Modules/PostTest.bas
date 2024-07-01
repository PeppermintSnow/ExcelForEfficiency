Attribute VB_Name = "PostTest"
' Module which handles the PostTest logic and events.

' Declare variables to keep track of scores.
Dim ScoreCorrect, ScoreIncorrect, ScoreGrade As Integer
Dim LastQuestion As Long

' Initializes the elements in the posttest.
Sub Initialize()
    ' Sets the variables to 0.
    Set Score = ActivePresentation.Slides(SlidePostResults) 'Posttest results screen slide
    ScoreCorrect = 0
    ScoreIncorrect = 0
    ScoreGrade = 0
    
    ' Sets the results screen text to 0.
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = ScoreGrade & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPost").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPost").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePost").TextFrame.TextRange = ScoreGrade & "%"
    
    ' Resets the boss' health bar text.
    ' Rounds off the integer to display with no decimals.
    HPText = Round((15 - ScoreCorrect) / 15 * 100, 0)
    ' Loop for slides with the boss health bar.
    For i = SlidePostFQ To SlidePostLast
        ' Loop for the 15 blocks in the boss' health bar.
        For n = 1 To 15
            ' Sets all of the boss' health visibility to true, and the color to green.
            ActivePresentation.Slides(i).Shapes("!!HPBar" & n).Visible = True
            ActivePresentation.Slides(i).Shapes("!!HPBar" & n).Fill _
            .ForeColor.RGB = RGB(141, 193, 21)
        Next n
            ' Displays the boss health count.
            ActivePresentation.Slides(i).Shapes("!!HPText").TextFrame.TextRange = HPText & "/100"
    Next i
    
    ' Calls subroutines to randomize the order of answers and to shuffle questions.
    RandomizeAnswerOrder
    ShuffleSlides
End Sub

' Registers a correct answer to the score count.
Sub CorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePostResults)
    ' Adds a correct point to the score count.
    ScoreCorrect = (ScoreCorrect) + 1
    ' Displays the correct score on their respective elements. 
    Score.Shapes("!!BoxCorrect").TextFrame.TextRange = ScoreCorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxCorrectPost").TextFrame.TextRange = ScoreCorrect
    ' Calls the OverallGrade subroutine to calculate score percentage.
    OverallGrade
    ' Calls the HPBarFunction subroutine to display the updated health count.
    HPBarFunction
End Sub

' Registers an incorrect answer to the score count.
Sub IncorrectAnswer()
    Set Score = ActivePresentation.Slides(SlidePostResults)
    ' Adds an incorrect point to the score count.
    ScoreIncorrect = (ScoreIncorrect) + 1
    ' Displays the incorrect score on their respective elements. 
    Score.Shapes("!!BoxIncorrect").TextFrame.TextRange = ScoreIncorrect
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxIncorrectPost").TextFrame.TextRange = ScoreIncorrect
    ' Calls the OverallGrade subroutine to calculate score percentage.
    OverallGrade
End Sub

' Calculates the score percentage.
Sub OverallGrade()
    Set Score = ActivePresentation.Slides(SlidePostResults)
    ' Formula for the Total score variable.
    Total = ScoreCorrect + ScoreIncorrect
    ' Formula for computing the percentage.
    ScoreGrade = Round(ScoreCorrect / Total * 100, 1)
    ' Displays the score percentage on their respective elements.
    Score.Shapes("!!BoxGrade").TextFrame.TextRange = (ScoreGrade) & "%"
    Score.Shapes("!!VBoxGrade").TextFrame.TextRange = ScoreGrade
    ActivePresentation.Slides(SlideFinalResults).Shapes("!!BoxGradePost").TextFrame.TextRange = ScoreGrade & "%"
End Sub

' Updates the boss' health bar.
Sub HPBarFunction()
    ' Declare a damage variable to subtract to health bar.
    Dim Damage As Integer
    ' Set the Damage variable to be equal to ScoreCorrect.
    Damage = ScoreCorrect
    ' Displays the health text by subtracting 15 to the correct score count, then converting it into a percentage with no decimals.
    HPText = Round((15 - ScoreCorrect) / 15 * 100, 0)
        ' Handles errors.
        On Error GoTo HandleError
        ' Loop to select and run the blocks of code on all slides with the health bar.
        For i = SlidePostFQ To SlidePostLast
            ' Sets the a part of the boss' health point's visibility to false based on the correct score count.
            ActivePresentation.Slides(i).Shapes("!!HPBar" & (Damage)).Visible = False
            ' Updates health count.
            ActivePresentation.Slides(i).Shapes("!!HPText").TextFrame.TextRange = HPText & "/100"
            ' Changes the health bar cololr if health is less than 8.
            If ActivePresentation.Slides(i).Shapes("!!HPBar8").Visible = False Then
                For n = 1 To 15
                    ActivePresentation.Slides(i).Shapes("!!HPBar" & (n)).Fill _
                    .ForeColor.RGB = RGB(229, 101, 46)
                Next n
            End If

            ' Changes the health bar cololr if health is less than 11.
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

' Shuffles slides.
Sub ShuffleSlides()
    FirstSlide = SlidePostFQ
    LastSlide = SlidePostLQ
    
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
    For i = SlidePostFQ To SlidePostLQ 
    
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

