Attribute VB_Name = "UploadResults"
' Responsible for handling score files. (Exports and Uploads)\

' Declare variables to keep track of scores.
Dim Correct, Incorrect, Grade, Name, Correct1, Incorrect1, Grade1 As Integer
Dim XCorrect, XIncorrect, XGrade, X1Correct, X1Incorrect, X1Grade As Variant

' Gets the username from the user.
Sub GetName()
    Dim UserName As String
    ' Displays an input box prompting for a name.
    UserName = InputBox("Please enter your name.", "UserName", "Enter your name here")
    ActivePresentation.SlideShowWindow.View.Next
    ActivePresentation.Slides(17).Shapes("!!Dialogue17").TextFrame.TextRange = ("Oh! Yes, my name is " & UserName & "!")
End Sub

' Calls subroutines for uploading results.
Sub UploadButton()
    GenerateBackupResults
    SendResultsToExcel
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Calls subroutines for uploading final results.
Sub finalUploadButton()
    GenerateAllBackupResults
    SendUSQToExcel
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Exports the results to an Excel workbook, which then gets uploaded to Google forms.
Sub SendResultsToExcel()
    ' Handles errors.
    On Error GoTo ErrorHandler
        ' Calls the GenerateBackupResults subroutine before uploading to forms.
        GenerateBackupResults

        ' Declares variables to run Excel.
        Dim FileName As String
    
        Dim xlsApp As Object
        Set xlsApp = CreateObject("Excel.Application")
        
        Dim xlsWB As Object
        Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\DATA.xlsm")
            ' Sets the test type to PreAssessment if the Pretest Checkpoint is false, else set to PostAssessment.
            If CheckpointPretest = False Then
                Set Results = ActivePresentation.Slides(SlidePreResults)
                TestType = "PreAssessment"
            Else
                Set Results = ActivePresentation.Slides(SlidePostResults)
                TestType = "PostAssessment"
            End If
                
                ' Stores the scores in a temporary variable.
                XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
                XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
                XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange

                ' Convert the strings stored in temporary variables to integers.
                Correct = CInt(XCorrect)
                Incorrect = CInt(XIncorrect)
                Grade = CInt(XGrade)

                ' Define cell addresses and values for the Excel workbook.
                xlsWB.Worksheets(1).Range("A1") = "Name"
                xlsWB.Worksheets(1).Range("A2") = UserName
                
                xlsWB.Worksheets(1).Range("B1") = "Correct"
                xlsWB.Worksheets(1).Range("B2") = Correct
                
                xlsWB.Worksheets(1).Range("C1") = "Incorrect"
                xlsWB.Worksheets(1).Range("C2") = Incorrect
                
                xlsWB.Worksheets(1).Range("D1") = "Overall Grade"
                xlsWB.Worksheets(1).Range("D2") = Grade
                
                xlsWB.Worksheets(1).Range("E1") = "Type"
                xlsWB.Worksheets(1).Range("E2") = TestType
                
                xlsWB.Worksheets(1).Range("J1") = "Send"
                
        ' Close Excel without saving. (Saving causes bugs for some reason)
        ' A block of code within the workbook will automatically run upon detecting changes and be uploaded to the Google forms.
        xlsWB.Close False
        xlsApp.Quit
        
        Set xlsWB = Nothing
        Set xlsApp = Nothing
    Exit Sub

' Notifies the user that an error has occured.
ErrorHandler:
    MsgBox "Error uploading score to database; a backup file was created.", vbCritical, "Error!"
    Resume Next
End Sub

' Exports the user satisfaction questionnaire to an Excel workbook.
Sub SendUSQToExcel()
        ' Handles errors.
        On Error GoTo ErrorHandler
        ' Generates backup results.
        GenerateBackupResults

        ' Declares variables for running Excel.
        Dim FileName As String
    
        Dim xlsApp As Object
        Set xlsApp = CreateObject("Excel.Application")
        
        Dim xlsWB As Object
        Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\usqDATA.xlsm")
                ' Sets the cell addresses and values.
                xlsWB.Worksheets(1).Range("A1") = "Q1"
                xlsWB.Worksheets(1).Range("A2") = USQ1
                
                xlsWB.Worksheets(1).Range("B1") = "Q2"
                xlsWB.Worksheets(1).Range("B2") = USQ2
                
                xlsWB.Worksheets(1).Range("C1") = "Q3"
                xlsWB.Worksheets(1).Range("C2") = USQ3
                
                xlsWB.Worksheets(1).Range("D1") = "Q4"
                xlsWB.Worksheets(1).Range("D2") = USQ4
                
                xlsWB.Worksheets(1).Range("E1") = "Q5"
                xlsWB.Worksheets(1).Range("E2") = USQ5
                
                xlsWB.Worksheets(1).Range("F1") = "Q6"
                xlsWB.Worksheets(1).Range("F2") = USQ6
                
                xlsWB.Worksheets(1).Range("G1") = "Q7"
                xlsWB.Worksheets(1).Range("G2") = USQ7
                
                xlsWB.Worksheets(1).Range("H1") = "Q8"
                xlsWB.Worksheets(1).Range("H2") = USQ8
                
                xlsWB.Worksheets(1).Range("J1") = "Send"

        ' Close without saving.
        xlsWB.Close False
        xlsApp.Quit
        
        Set xlsWB = Nothing
        Set xlsApp = Nothing
    Exit Sub

' Notifies the user that an error has occured.
ErrorHandler:
    MsgBox "Error uploading score to database; a backup file was created.", vbCritical, "Error!"
    Resume Next
End Sub

' Generates backup results. (.txt file)
Sub GenerateBackupResults()
    ' If pretest checkpoint is false, test type will be PreAssessment, otherwise, PostAssessment.
    If CheckpointPretest = False Then
        Set Results = ActivePresentation.Slides(SlidePreResults)
        TestType = "PreAssessment"
    Else
        Set Results = ActivePresentation.Slides(SlidePostResults)
        TestType = "PostAssessment"
    End If
    
    ' Stores the scores in a temporary variable.
    XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
    XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
    XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange

    ' Convert strings to integers.
    Correct = CInt(XCorrect)
    Incorrect = CInt(XIncorrect)
    Grade = CInt(XGrade)
    
    ' Define variables to export to .txt
    Dim txtPath As String

    ' Set target path.
    txtPath = (ActivePresentation.Path & "\" & TestType & "DATA.txt")
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim txtFile As TextStream
    ' Define values for the lines of the text file.
    Set txtFile = FSO.CreateTextFile(txtPath, True)
        txtFile.WriteLine ("Name=" & UserName)
        txtFile.WriteLine ("Correct=" & Correct)
        txtFile.WriteLine ("Incorrect=" & Incorrect)
        txtFile.WriteLine ("Grade=" & Grade)
        txtFile.WriteLine ("Type=" & TestType)
    txtFile.Close
    
    Set txtFile = Nothing
    Set FSO = Nothing
End Sub

' Generates a collective final backup results for both pretest, posttest, and user satisfaction questionnaire.
Sub GenerateAllBackupResults()
    ' Sets the target slides to gather scores from.
    Set Results = ActivePresentation.Slides(SlidePreResults)
    TestType = "PreAssessment"

    Set Results1 = ActivePresentation.Slides(SlidePostResults)
    TestType = "PostAssessment"
    
    ' Stores scores in temporary variables.
    XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
    XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
    XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange
    
    X1Correct = Results1.Shapes("!!BoxCorrect").TextFrame.TextRange
    X1Incorrect = Results1.Shapes("!!BoxIncorrect").TextFrame.TextRange
    X1Grade = Results1.Shapes("!!VBoxGrade").TextFrame.TextRange

    ' Converts strings to integers.
    Correct = CInt(XCorrect)
    Incorrect = CInt(XIncorrect)
    Grade = CInt(XGrade)
    
    Correct1 = CInt(X1Correct)
    Incorrect1 = CInt(X1Incorrect)
    Grade1 = CInt(X1Grade)
    
    ' Declare variables to export to .txt
    Dim txtPath As String
    ' Sets target path.
    txtPath = (ActivePresentation.Path & "\" & "finalDATA.txt")
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim txtFile As TextStream
    ' Define values for the lines in the text file.
    Set txtFile = FSO.CreateTextFile(txtPath, True)
        txtFile.WriteLine ("Name = " & UserName)
        txtFile.WriteLine ("Correct = " & Correct)
        txtFile.WriteLine ("Incorrect = " & Incorrect)
        txtFile.WriteLine ("Grade = " & Grade)
        txtFile.WriteLine ("Type = " & "PreTest")
        txtFile.WriteBlankLines (1)
        txtFile.WriteLine ("Correct = " & Correct1)
        txtFile.WriteLine ("Incorrect = " & Incorrect1)
        txtFile.WriteLine ("Grade = " & Grade1)
        txtFile.WriteLine ("Type = " & "PostTest")
        txtFile.WriteBlankLines (1)
        txtFile.WriteLine ("USQ1 = " & USQ1)
        txtFile.WriteLine ("USQ2 = " & USQ2)
        txtFile.WriteLine ("USQ3 = " & USQ3)
        txtFile.WriteLine ("USQ4 = " & USQ4)
        txtFile.WriteLine ("USQ5 = " & USQ5)
        txtFile.WriteLine ("USQ6 = " & USQ6)
        txtFile.WriteLine ("USQ7 = " & USQ7)
        txtFile.WriteLine ("USQ8 = " & USQ8)
    txtFile.Close
    
    Set txtFile = Nothing
    Set FSO = Nothing
    ActivePresentation.SlideShowWindow.View.Next
End Sub
