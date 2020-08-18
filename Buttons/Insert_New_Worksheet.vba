Sub Insert_New_Worksheet()
    ' insert new sheet
    
    Worksheets.Add
    
    ' Add Values
    
    Range("A1").Value = "Hello World"
    Range("A2").Value = Date
    Range("A3").Value = Time
    
    ' Format Cells
    Range("A1:A3").Borders.LineStyle = xlContinuous
    
    Columns("A").AutoFit
End Sub
