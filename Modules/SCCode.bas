Attribute VB_Name = "SCCode"
' Module responsible for handling events in Sorting and Charts (Aurora).

' Checks if the sample activity has been accomplished.
Sub ResponseFinishSorting()
    If SampleCheck = True Then
        ActivePresentation.SlideShowWindow.View.Next
    Else
        MsgBox "You have not yet accomplished the task; please go back to the sample worksheet and finish the task.", vbExclamation, ":<"
    End If
    
    SampleCheck = False
End Sub

' Opens the sample activity.
Sub SampleOpenSorting()
    SampleCheck = True
    
    Dim xlsApp As Object
    Set xlsApp = CreateObject("Excel.Application")
    
    Dim xlsWB As Object
    Set xlsWB = xlsApp.Workbooks.Open(ActivePresentation.Path & "\SampleSorting.xlsm")
    
    xlsApp.Visible = True
    
    Set xlsApp = Nothing
    Set xlsWB = Nothing
End Sub
