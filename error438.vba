Sub OpenExcelFile()
    Dim PowerPointApp As Object
    Dim PowerPointPres As Object
    Dim ExcelApp As Object
    Dim ExcelWB As Object
    Dim FilePath As String
    
    ' Create an instance of PowerPoint
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    
    ' Create a new presentation
    Set PowerPointPres = PowerPointApp.Presentations.Add
    
    ' Get the folder path of the PowerPoint presentation
    FilePath = Left(PowerPointPres.Path, InStrRev(PowerPointPres.Path, "\")) & "book.xls"
    
    ' Create an instance of Excel
    Set ExcelApp = CreateObject("Excel.Application")
    
    ' Open the Excel file
    Set ExcelWB = ExcelApp.Workbooks.Open(FilePath)
    
    ' Show Excel application
    ExcelApp.Visible = True
    
    ' Clean up objects
    Set ExcelWB = Nothing
    Set ExcelApp = Nothing
    
    ' Release the PowerPoint objects
    Set PowerPointPres = Nothing
    Set PowerPointApp = Nothing
End Sub

