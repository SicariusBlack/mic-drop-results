Attribute VB_Name = "Module1"
Function DelSlide(index As Integer)
    ActivePresentation.Slides(index).Delete
End Function

Function Count()
    Count = ActivePresentation.Slides.Count
End Function

Function Duplicate(index As Integer)
    Dim oSlide As Slide
    Set oSlide = ActivePresentation.Slides(index).Duplicate()(1)
    oSlide.MoveTo toPos:=ActivePresentation.Slides.Count
End Function

Function SaveAs(filename As String)
    ActivePresentation.SaveAs (filename)
End Function
