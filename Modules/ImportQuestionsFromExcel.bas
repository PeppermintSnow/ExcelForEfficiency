Attribute VB_Name = "ImportQuestionsFromExcel"
' Module responsible for the automation of importing the test questions from an Excel workbook.
' Only used during development.

Sub ImportQuestions()
    ' Declare and set variable for opening the Excel workbook.
    Dim xlsWB As Object
    Set xlsWB = CreateObject("Excel.Application").Workbooks.Open(ActivePresentation.Path & "\QuestionSheet.xlsx")
    
        For i = 1 To 2 'Number of slides.
            ActivePresentation.Slides(i).Shapes("!!QuestionBox").TextFrame.TextRange = xlsWB.Worksheets(2).Range("A" & i - 51)
            ActivePresentation.Slides(i).Shapes("!!ChoiceA").TextFrame.TextRange = xlsWB.Worksheets(2).Range("B" & i - 51)
            ActivePresentation.Slides(i).Shapes("!!ChoiceB").TextFrame.TextRange = xlsWB.Worksheets(2).Range("C" & i - 51)
            ActivePresentation.Slides(i).Shapes("!!ChoiceC").TextFrame.TextRange = xlsWB.Worksheets(2).Range("D" & i - 51)
            ActivePresentation.Slides(i).Shapes("!!ChoiceD").TextFrame.TextRange = xlsWB.Worksheets(2).Range("E" & i - 51)
        Next i
        'Integer i must be equal to 2; remember to adjust the minus value later on when adjusting the slide number.
        
    ' Close the workbook.
    xlsWB.Close
    Set xlsWB = Nothing
End Sub
