Sub CopyTable()

    'Copy Contact list and paste to new sheet
    
    Worksheets("Contact List").Select
    
    Range("A1:C51").Copy Destination:=Worksheets("Sheet1").Range("A1")
    
    'Autofit columns
    Worksheets("Sheet1").Select
    Columns("A:C").AutoFit
    
    
    
End Sub
