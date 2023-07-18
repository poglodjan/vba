Sub Example()
    Dim myChart As Object
    
    ' Call the function to display the chart in a slide
    Set myChart = DisplayChartInSlide()
    
    ' Modify the chart properties if needed
    myChart.ChartTitle.Text = "Sales Data"
    myChart.Axes(xlValue).HasTitle = True
    myChart.Axes(xlValue).AxisTitle.Text = "Amount"
    
    ' ... additional chart manipulation code ...
    
    ' Clean up the chart object
    Set myChart = Nothing
End Sub

Function DisplayChartInSlide() As Object
    Dim PowerPointApp As Object
    Dim PowerPointPres As Object
    Dim PowerPointSlide As Object
    Dim ChartObj As Object
    Dim ChartData As Object
    
    ' Create an instance of PowerPoint
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    ' Create a new presentation
    Set PowerPointPres = PowerPointApp.Presentations.Add
    
    ' Add a slide to the presentation
    Set PowerPointSlide = PowerPointPres.Slides.Add(1, 12) ' 12 represents the slide layout type
    
    ' Add a chart to the slide
    Set ChartObj = PowerPointSlide.Shapes.AddChart(xlColumnClustered, 100, 100, 400, 300) ' Adjust the position and size as per your requirement
    
    ' Set the chart data
    Set ChartData = ChartObj.Chart
    ChartData.ChartData.Activate
    ChartData.ChartData.Workbook.Sheets(1).Cells(1, 1).Value = "Category"
    ChartData.ChartData.Workbook.Sheets(1).Cells(1, 2).Value = "Value 1"
    ChartData.ChartData.Workbook.Sheets(1).Cells(2, 1).Value = "A"
    ChartData.ChartData.Workbook.Sheets(1).Cells(2, 2).Value = 10
    ChartData.ChartData.Workbook.Sheets(1).Cells(3, 1).Value = "B"
    ChartData.ChartData.Workbook.Sheets(1).Cells(3, 2).Value = 20
    
    ' Update the chart data source range
    ChartData.SetSourceData Source:=ChartData.ChartData.Workbook.Sheets(1).Range("A1:B3")
    
    ' Show PowerPoint application
    PowerPointApp.Visible = True
    
    ' Set the return value to the chart object
    Set DisplayChartInSlide = ChartObj
    
    ' Clean up objects
    Set ChartData = Nothing
    Set ChartObj = Nothing
    Set PowerPointSlide = Nothing
    
    ' Release the PowerPoint objects
    Set PowerPointPres = Nothing
    Set PowerPointApp = Nothing
End Function
