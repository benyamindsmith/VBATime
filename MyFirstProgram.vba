Sub MyFirstProgram()

    ' Create a worksheet
    Worksheets.Add
    
    ' Enter values into cells
    Range("A1").Value = "Hello World!"
    Range("A2").Value = Date
    
    ' Format Cells
    Range("A1:A2").Interior.Color = rgbCoral
    
    
    
End Sub
