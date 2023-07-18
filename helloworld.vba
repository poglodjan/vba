Sub DisplayStringOnSlide()
    Dim PowerPointApp As Object
    Dim PowerPointPres As Object
    Dim PowerPointSlide As Object
    Dim MyString As String
    
    ' Declare the string variable
    MyString = "Hello world"
    
    ' Create an instance of PowerPoint
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    ' Create a new presentation
    Set PowerPointPres = PowerPointApp.Presentations.Add
    
    ' Add a slide to the presentation
    Set PowerPointSlide = PowerPointPres.Slides.Add(1, 12) ' 12 represents the slide layout type
    
    ' Add a text box to the slide and display the string
    PowerPointSlide.Shapes.AddTextbox 1, 100, 100, 400, 100 ' Adjust the position and size as per your requirement
    PowerPointSlide.Shapes(1).TextFrame.TextRange.Text = MyString
    
    ' Show PowerPoint application
    PowerPointApp.Visible = True
    
    ' Clean up objects
    Set PowerPointSlide = Nothing
    
    ' Release the PowerPoint objects
    Set PowerPointPres = Nothing
    Set PowerPointApp = Nothing
End Sub
