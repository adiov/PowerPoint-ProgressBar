  Sub AddProgressBar()
    On Error Resume Next
        With ActivePresentation
              For X = 2 To .Slides.Count
              .Slides(X).Shapes("PB_BG").Delete
              .Slides(X).Shapes("PB_PB").Delete
              Set progressBackground = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
              0, .PageSetup.SlideHeight - 12, .PageSetup.SlideWidth, 12)
              progressBackground.Fill.Solid.ForeColor.RGB = RGB(255, 255, 255)
              progressBackground.Name = "PB_BG"
              Set progressBar = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
              0, .PageSetup.SlideHeight - 12, _
              X * .PageSetup.SlideWidth / .Slides.Count, 12)
              progressBar.Fill.Solid.ForeColor.RGB = RGB(0, 130, 200)
              progressBar.Name = "PB_PB"
              .Slides(X).Shapes("PB_BG").Fill.ForeColor.RGB = RGB(255, 255, 255)
              .Slides(X).Shapes("PB_BG").Line.Visible = msoFalse
              .Slides(X).Shapes("PB_PB").Fill.ForeColor.RGB = RGB(0, 130, 200)
              .Slides(X).Shapes("PB_PB").Line.Visible = msoFalse
              .Slides(X).Shapes("PB_PB").TextFrame.TextRange.Text = Round(X / .Slides.Count * 100) & "%"
              .Slides(X).Shapes("PB_PB").TextFrame.TextRange.Font.Size = 10
              .Slides(X).Shapes("PB_PB").TextFrame.TextRange.Font.Bold = msoCTrue
              .Slides(X).Shapes("PB_PB").TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignRight
              Next X:
        End With
End Sub

