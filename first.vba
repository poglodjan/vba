Sub CreateChart()

    Dim ws As Worksheet
    Dim ch As Chart
    Dim dt As Range
    
    Set ws = ActiveSheet
    Set dt = Range("A2:D9")
    Set ch = ws.Shapes.AddChart2(Style:=280, Width:=300, Height:=300, Left:=Range("F1").Left, Top:=Range("F1").Top).Chart
    
    With ch
        .SetSourceData Source:=dt
        .ChartType = xlBarClustered
        .ChartTitle.Text = "ye"
        .SetElement msoElementDataLabelOutSideEnd
        .SetElement msoElementLegendTop
        .SetElement msoElementPrimaryValueAxisNone
        .SetElement msoElementPrimaryCategoryAxisTitleBelowAxis
        .Axes(xlCategory).AxisTitle.Text = "tegion"
        .SeriesCollection("Serie1").Interior.Color = RGB(250, 0, 0)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(221, 227, 185)
        
    End With
End Sub


Sub AddCharts()

Dim i As Integer 'rows
Dim j As Integer 'columns

i = Cells(Rows.Count, 1).End(xlUp).Row

For j = 2 To 4
    With ActiveSheet.Shapes.AddChart.Chart
    .ChartType = xlXYScatter
    .SeriesCollection.NewSeries
        With .SeriesCollection(1)
        .Name = "=" & ActiveSheet.Name & "!" & _
        Cells(1, j).Address
        .XValues = "=" & ActiveSheet.Name & "!" & _
        Range(Cells(2, 1), Cells(i, 1)).Address
        .Values = "=" & ActiveSheet.Name & "!" & _
        Range(Cells(2, j), Cells(i, j)).Address
        End With
    .HasLegend = False
    End With
Next j

End Sub

Sub WykresLiniaZacienienie()
    Dim chartObject As chartObject
    Dim chartData As Range
    Dim x As Double
    Dim y As Double
    
    ' Utwórz nowy obiekt wykresu
    Set chartObject = ActiveSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=75, Height:=225)
    
    ' Utwórz dane dla wykresu y=x
    For x = -10 To 10
        y = x
        Cells(x + 11, 1).Value = x
        Cells(x + 11, 2).Value = y
    Next x
    
    ' Ustaw zakres danych dla wykresu
    Set chartData = Range("A1:B21")
    
    ' Dodaj dane do wykresu
    With chartObject.Chart
        .SetSourceData Source:=chartData
        .ChartType = xlLine
        
        ' Zaznacz pole pod linią y=x
        Dim fillColor As Long
        fillColor = RGB(192, 192, 192) ' Kolor szary
        
        For x = -10 To 10
            y = x
            .SeriesCollection(1).Points(x + 11).Format.Fill.ForeColor.RGB = fillColor
        Next x
    End With
End Sub


