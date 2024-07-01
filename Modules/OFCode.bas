Attribute VB_Name = "OFCode"
' Module responsible for handling events in the Operations and Functions section (Xenolumina).

' Sets Checkpoint of Xenolumina's intro to true.
Sub TickCheckpointOFIntro()
    CheckpointOFIntro = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Checks if the intro has been played.
Sub PlanetEntrance()
    If CheckpointOFIntro = False Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideOFMenu
End Sub

' Checks if the user has accomplished the Operations activity.
Sub ResponseFinishOFO()
    ' Allows the user to proceed if the sample activity variable is set to true, otherwise, displays an error message.
    If SampleCheck = True Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "You have not yet accomplished the task; please go back to the sample worksheet and finish the task.", vbExclamation, ":<"
    End If
    
    ' Reset the sample activity variable to false for future use.
    SampleCheck = False
End Sub

' Checks if the user has accomplished the Functions activity.
Sub ResponseFinishOFF()
    ' Allows the user to proceed if the sample activity variable is set to true, otherwise, displays an error message.
    If SampleCheck = True Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "You have not yet accomplished the task; please go back to the sample worksheet and finish the task.", vbExclamation, ":<"
    End If
    
    ' Reset the sample activity variable to false for future use.
    SampleCheck = False
End Sub

' Opens the sample workbook activity for Operations.
Sub SampleOpenOFO()
    ' Set sample activity variable to true upon launch.
    SampleCheck = True
    
    ' Declare variables for opening the workbook.
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\SampleOperations.xlsm")
    
    ' Makes the Excel workbook visible.
    xlsApp.Visible = True
    
    ' Resets the variables to avoid conflicts.
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub

' Opens the sample workbook activity for Operations.
Sub SampleOpenOFF()
    ' Set sample activity variable to true upon launch.
    SampleCheck = True
    
    ' Declare variables for opening the workbook.
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\SampleFunctions.xlsm")
    
    ' Makes the Excel workbook visible.
    xlsApp.Visible = True
    
    ' Resets the variables to avoid conflicts.
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub

' Closes the Excel workbook. (Unused)
Sub ExcelClose()
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\OperationsSample.xlsm")
    
    xlsWB.Close False
    xlsApp.Quit
    
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub

' Handles the textbox activity for slide 196.
Sub ButtonIFGo()
    ' Declare variable for checking the string.
    Dim UserInput As String
    UserInput = Slide196.UserInput.Value

    ' Checks if the user's string contains "Average"
    If InStr(1, UserInput, "Average", vbTextCompare) Then
        Slide196.UserInput.Value = ""
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "Please type [Average] on the textbox.", vbExclamation, "Error!"
    End If
End Sub

' Handles the textbox activity for slide 199. 
Sub ButtonIFGo2()
    ' Declare variable for checking the string.
    Dim UserInput As String
    UserInput = Slide199.UserInput.Value

    ' Checks if the user's string contains "B1:B4"
    If UserInput = "B1:B4" Then
        Slide199.UserInput.Value = ""
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "Please type [B1:B4] on the textbox.", vbExclamation, "Error!"
    End If
End Sub
