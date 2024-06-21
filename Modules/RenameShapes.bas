Attribute VB_Name = "RenameShapes"
Sub Rename()

    On Error GoTo HandleError
        For i = 1 To 303
                ActivePresentation.Slides(i).Shapes("!!Dialogue").Name = "!!Dialogue" & i
        Next i
        
    Exit Sub
        
HandleError:
    Resume Next
End Sub

Sub Resize()
    On Error GoTo HandleError
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
Sub AddAction()
    On Error GoTo HandleError
        For i = 35 To 49
            For n = 2 To 4
            Set Choice = ActivePresentation.Slides(i).Shapes("Choice" & n)
            
            With Choice.ActionSettings(ppMouseClick)
                .Action = ppActionRunMacro
                .Run = "PreTest.IncorrectAnswer"
            End With
            Next n
                
        Next i
        
    Exit Sub
        
HandleError:
    Resume Next
End Sub

Sub ChangeText()
    On Error GoTo HandleError
        For i = 52 To 170
            Set Slide = ActivePresentation.Slides(i)
                Slide.Shapes("!!LabelSC").TextFrame.TextRange = "Aurora"
                Slide.Shapes("!!LabelAS").TextFrame.TextRange = "Tenebris"
                Slide.Shapes("!!LabelOF").TextFrame.TextRange = "Xenolumina"
        Next i
    Exit Sub
    
HandleError:
    Resume Next
            
End Sub
Sub ChangeShapeProperty()
    On Error GoTo HandleError
        For i = 6 To 23
                With ActivePresentation.Slides(i).Shapes("LabelAssessment")
                    .TextFrame.TextRange.Font.Color.RGB = RGB(180, 139, 234)
                    .Glow.Color.RGB = RGB(61, 45, 91)
                    .Glow.Radius = 10
                    .Line.Visible = msoFalse
                End With
        Next i
    Exit Sub
HandleError:
    Resume Next
End Sub

Sub ChangeTextAnimation()
    On Error GoTo HandleError
        For i = 35 To 36
            With ActivePresentation.Slides(i).Shapes("!!Dialogue").AnimationSettings
                
End Sub

Sub DeleteShape()
    On Error GoTo HandleError
        For i = 256 To 287
            With ActivePresentation.Slides(i)
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

Sub SendToBack()
    On Error GoTo HandleError
        For i = 32 To 55
            With ActivePresentation.Slides(i)
                .Shapes("!!TransitionTop").ZOrder msoBringToFront
                .Shapes("!!TransitionBot").ZOrder msoBringToFront
            End With
        Next i
    Exit Sub
HandleError:
    Resume Next
End Sub

Sub OnOddNumber()
    On Error GoTo HandleError
        For i = 1 To 303
            With ActivePresentation.Slides(i)
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

Sub AddSFX()
    On Error GoTo HandleError
        For i = 1 To 303
            With ActivePresentation.Slides(i).Shapes("!!Choice4").ActionSettings(ppMouseClick)
                .SoundEffect.ImportFromFile "C:\Users\Joa\Downloads\wrong2.wav"
            End With
        Next i
    Exit Sub
HandleError:
    Resume Next
End Sub

