Attribute VB_Name = "RenameShapes"
' Module for automating the process of renaming shapes and modifying several properties, too.
' Only used during development; not used in running the program.

' Renames shapes.
Sub Rename()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Changes the name of a shape from "X" to "Y" in the given slide numbers.
        For i = 1 To 303
                ActivePresentation.Slides(i).Shapes("!!Dialogue").Name = "!!Dialogue" & i
        Next i
    Exit Sub
        
HandleError:
    Resume Next
End Sub

' Resizes shapes.
Sub Resize()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Changes the height and width dimensions of the chosen shapes in the given slide numbers.
        For i = 74 To 97
            ActivePresentation.Slides(i).Shapes("!!Choice1").Height = 70.16835
            ActivePresentation.Slides(i).Shapes("!!Choice1").Width = 190.8
            
            ActivePresentation.Slides(i).Shapes("!!Choice2").Height = 70.16835
            ActivePresentation.Slides(i).Shapes("!!Choice2").Width = 190.8
            
            ActivePresentation.Slides(i).Shapes("!!Choice3").Height = 70.16835
            ActivePresentation.Slides(i).Shapes("!!Choice3").Width = 190.8
            
            ActivePresentation.Slides(i).Shapes("!!Choice4").Height = 70.16835
            ActivePresentation.Slides(i).Shapes("!!Choice4").Width = 190.8
            
        Next i
        
    Exit Sub
            
HandleError:
    Resume Next
End Sub

' Adds an action to the shapes.
Sub AddAction()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Selects the slide numbers.
        For i = 35 To 49
            ' (Optional for multiple selections) Run a loop.
            For n = 2 To 4
            ' Set target shape.
            Set Choice = ActivePresentation.Slides(i).Shapes("Choice" & n)
            
            ' Modify action properties.
            With Choice.ActionSettings(ppMouseClick) ' Detects when clicked.
                .Action = ppActionRunMacro ' Specifies the action to run, in this case, it runs a macro.
                .Run = "PreTest.IncorrectAnswer" ' Specifies which macro to run.
            End With
            Next n
                
        Next i
        
    Exit Sub
        
HandleError:
    Resume Next
End Sub

' Changes the text displayed.
Sub ChangeText()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Set target slides.
        For i = 52 To 170
            Set Slide = ActivePresentation.Slides(i)
                ' Changes a shape's text from "X" to "Y"
                Slide.Shapes("!!LabelSC").TextFrame.TextRange = "Aurora"
                Slide.Shapes("!!LabelAS").TextFrame.TextRange = "Tenebris"
                Slide.Shapes("!!LabelOF").TextFrame.TextRange = "Xenolumina"
        Next i
    Exit Sub
    
HandleError:
    Resume Next
End Sub

' Deletes shapes.
Sub DeleteShape()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Sets target slides.
        For i = 256 To 287
            With ActivePresentation.Slides(i)
                ' Deletes target shapes.
                .Shapes("!!PlanetSurface").Delete
                .Shapes("!!BGSpace").Delete
                .Shapes("!!BossShadow").Delete
                .Shapes("!!BobShadow").Delete
            End With
        Next i
    Exit Sub

HandleError:
    Resume Next
End Sub

' Modifies shape z-index order. (Sets z-index)
Sub SendToBack()
    ' Handles errors in case a shape is not detected in a slide.
    On Error GoTo HandleError
        ' Sets target slides.
        For i = 32 To 55
            With ActivePresentation.Slides(i)
                ' Brings target to front.
                .Shapes("!!TransitionTop").ZOrder msoBringToFront
                .Shapes("!!TransitionBot").ZOrder msoBringToFront
            End With
        Next i
    Exit Sub

HandleError:
    Resume Next
End Sub

' Detects if Odd number.
Sub OnOddNumber()
    On Error GoTo HandleError
        ' Set target slides.
        For i = 1 To 303
            With ActivePresentation.Slides(i)
                ' Conditional to rename dialogues based on odd or even.
                If i Mod 2 = 0 Then
                    .Shapes("!!Dialogue").Name = "!!Dialogue"
                Else
                    .Shapes("!!Dialogue").Name = "!!Dialogue1"
                End If
            End With
        Next i
    Exit Sub

HandleError:
    Resume Next
End Sub

' Adds sound effects.
Sub AddSFX()
    ' Handles errors in case a shape is not detected in a slide.
        ' Sets target slides.
        For i = 1 To 303
            ' Imports and sets a sound effect to play on click.
            With ActivePresentation.Slides(i).Shapes("!!Choice4").ActionSettings(ppMouseClick)
                .SoundEffect.ImportFromFile "C:\Users\[REDACTED]\Downloads\wrong2.wav"
            End With
        Next i
    Exit Sub
    
HandleError:
    Resume Next
End Sub

