Sub AddSampleTableToSlide()
    Dim PowerPointApp As Object
    Dim PowerPointPres As Object
    Dim PowerPointSlide As Object
    Dim TableObj As Object
    Dim SlideWidth As Double
    Dim SlideHeight As Double
    Dim SampleTable As Range
    Dim i As Long, j As Long
    
    ' Create an instance of PowerPoint
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    ' Create a new presentation
    Set PowerPointPres = PowerPointApp.Presentations.Add
    
    ' Add a slide to the presentation
    Set PowerPointSlide = PowerPointPres.Slides.Add(1, 12) ' 12 represents the slide layout type
    
    ' Set up a sample table
    Set SampleTable = ThisWorkbook.Worksheets("Sheet1").Range("A1:D5")
    
    ' Set slide dimensions based on table size
    SlideWidth = SampleTable.Columns.Count * 100 ' Adjust the column width as needed
    SlideHeight = SampleTable.Rows.Count * 40 ' Adjust the row height as needed
    
    ' Set the table position and dimensions on the slide
    Set TableObj = PowerPointSlide.Shapes.AddTable(SampleTable.Rows.Count + 1, SampleTable.Columns.Count, 100, 100, SlideWidth, SlideHeight).Table
    
    ' Copy table data from the range to the PowerPoint table
    For i = 1 To SampleTable.Rows.Count
        For j = 1 To SampleTable.Columns.Count
            TableObj.Cell(i, j).Shape.TextFrame.TextRange.Text = SampleTable.Cells(i, j).Value
        Next j
    Next i
    
    ' Show PowerPoint application
    PowerPointApp.Visible = True
    
    ' Clean up objects
    Set TableObj = Nothing
    
    ' Release the PowerPoint objects
    Set PowerPointSlide = Nothing
    Set PowerPointPres = Nothing
    Set PowerPointApp = Nothing
End Sub
