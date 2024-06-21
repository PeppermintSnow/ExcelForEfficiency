Attribute VB_Name = "OFCode"
Sub TickCheckpointOFIntro()
    CheckpointOFIntro = True
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub PlanetEntrance()
    If CheckpointOFIntro = False Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        ActivePresentation.SlideShowWindow.View.GotoSlide SlideOFMenu
End Sub

Sub ResponseFinishOFO()
    If SampleCheck = True Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "You have not yet accomplished the task; please go back to the sample worksheet and finish the task.", vbExclamation, ":<"
    End If
    
    SampleCheck = False
End Sub

Sub ResponseFinishOFF()
    If SampleCheck = True Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "You have not yet accomplished the task; please go back to the sample worksheet and finish the task.", vbExclamation, ":<"
    End If
    
    SampleCheck = False
End Sub


Sub SampleOpenOFO()
    SampleCheck = True
    
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\SampleOperations.xlsm")
    
    xlsApp.Visible = True
    
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub

Sub SampleOpenOFF()
    SampleCheck = True
    
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\SampleFunctions.xlsm")
    
    xlsApp.Visible = True
    
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub


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

Sub ButtonIFGo()
    Dim UserInput As String
    UserInput = Slide196.UserInput.Value
    If InStr(1, UserInput, "Average", vbTextCompare) Then
        Slide196.UserInput.Value = ""
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "Please type [Average] on the textbox.", vbExclamation, "Error!"
    End If
    
End Sub

Sub ButtonIFGo2()
    Dim UserInput As String
    UserInput = Slide199.UserInput.Value
    If UserInput = "B1:B4" Then
        Slide199.UserInput.Value = ""
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "Please type [B1:B4] on the textbox.", vbExclamation, "Error!"
    End If
End Sub
