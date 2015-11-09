  Sub RemoveProgressBar()
    On Error Resume Next
        With ActivePresentation
              For X = 1 To .Slides.Count
              .Slides(X).Shapes("PB_BG").Delete
              .Slides(X).Shapes("PB_PB").Delete
              Next X:
        End With
End Sub
