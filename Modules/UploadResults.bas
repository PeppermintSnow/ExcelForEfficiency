Attribute VB_Name = "UploadResults"
Dim Correct, Incorrect, Grade, Name, Correct1, Incorrect1, Grade1 As Integer
Dim XCorrect, XIncorrect, XGrade, X1Correct, X1Incorrect, X1Grade As Variant

Sub GetName()
    Dim UserName As String
    UserName = InputBox("Please enter your name.", "UserName", "Enter your name here")
    ActivePresentation.SlideShowWindow.View.Next
    ActivePresentation.Slides(17).Shapes("!!Dialogue17").TextFrame.TextRange = ("Oh! Yes, my name is " & UserName & "!")
End Sub

Sub UploadButton()
    GenerateBackupResults
    SendResultsToExcel
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub finalUploadButton()
    GenerateAllBackupResults
    SendUSQToExcel
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub SendResultsToExcel()
        On Error GoTo ErrorHandler
        GenerateBackupResults
        Dim FileName As String
    
        Dim xlsApp As Object
        Set xlsApp = CreateObject("Excel.Application")
        
        Dim xlsWB As Object
        Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\DATA.xlsm")
        
            If CheckpointPretest = False Then
                Set Results = ActivePresentation.Slides(SlidePreResults)
                TestType = "PreAssessment"
            Else
                Set Results = ActivePresentation.Slides(SlidePostResults)
                TestType = "PostAssessment"
            End If
                
                XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
                XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
                XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange
    
                Correct = CInt(XCorrect)
                Incorrect = CInt(XIncorrect)
                Grade = CInt(XGrade)
        
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
                
        xlsWB.Close False
        xlsApp.Quit
        
        Set xlsWB = Nothing
        Set xlsApp = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Error uploading score to database; a backup file was created.", vbCritical, "Error!"
    Resume Next
End Sub

Sub SendUSQToExcel()
        On Error GoTo ErrorHandler
        GenerateBackupResults
        Dim FileName As String
    
        Dim xlsApp As Object
        Set xlsApp = CreateObject("Excel.Application")
        
        Dim xlsWB As Object
        Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\usqDATA.xlsm")
        
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
        xlsWB.Close False
        xlsApp.Quit
        
        Set xlsWB = Nothing
        Set xlsApp = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Error uploading score to database; a backup file was created.", vbCritical, "Error!"
    Resume Next
End Sub

Sub GenerateBackupResults()
    If CheckpointPretest = False Then
        Set Results = ActivePresentation.Slides(SlidePreResults)
        TestType = "PreAssessment"
    Else
        Set Results = ActivePresentation.Slides(SlidePostResults)
        TestType = "PostAssessment"
    End If
    
    XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
    XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
    XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange

    Correct = CInt(XCorrect)
    Incorrect = CInt(XIncorrect)
    Grade = CInt(XGrade)
    
    Dim txtPath As String
    txtPath = (ActivePresentation.Path & "\" & TestType & "DATA.txt")
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim txtFile As TextStream
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

Sub GenerateAllBackupResults()
    Set Results = ActivePresentation.Slides(SlidePreResults)
    TestType = "PreAssessment"

    Set Results1 = ActivePresentation.Slides(SlidePostResults)
    TestType = "PostAssessment"
    
    XCorrect = Results.Shapes("!!BoxCorrect").TextFrame.TextRange
    XIncorrect = Results.Shapes("!!BoxIncorrect").TextFrame.TextRange
    XGrade = Results.Shapes("!!VBoxGrade").TextFrame.TextRange
    
    X1Correct = Results1.Shapes("!!BoxCorrect").TextFrame.TextRange
    X1Incorrect = Results1.Shapes("!!BoxIncorrect").TextFrame.TextRange
    X1Grade = Results1.Shapes("!!VBoxGrade").TextFrame.TextRange

    Correct = CInt(XCorrect)
    Incorrect = CInt(XIncorrect)
    Grade = CInt(XGrade)
    
    Correct1 = CInt(X1Correct)
    Incorrect1 = CInt(X1Incorrect)
    Grade1 = CInt(X1Grade)
    
    Dim txtPath As String
    txtPath = (ActivePresentation.Path & "\" & "finalDATA.txt")
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim txtFile As TextStream
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

Sub ExcelClose()
    Sheet1 = "\DATA.xlsm"
    Sheet2 = "\SampleOperations.xlsm"
    Sheet3 = "\SampleFunctions.xlsm"
    
    On Error GoTo ErrorHandler
            Dim xlsApp As Object
            Set xlsApp = CreateObject("Excel.Application")
            
            Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & Sheet1)
        
            xlsWB.Close False
            xlsApp.Quit
        
            Set xlsApp = Nothing
            Set xlsWB = Nothing
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Sub test()
    UserCode = Slide271.UserInputNumber.Value
    Debug.Print UserCode
End Sub
