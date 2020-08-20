Option Explicit

Sub ExtractEmails()
    
    '####################################
    'Make sure to first go to Tools > References...
    'and check off "Microsoft VBScript Regular Expressions 5.5
    '#####################################
    
    ' Define Range to apply function- "B2:B51" is where the relevant data is
    Dim arr As Variant
    arr = Range("B2:B51")

        
    ' Create Regex object
    Dim regEx As New RegExp
    regEx.Global = True
    
    ' Define pattern
    regEx.Pattern = "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)"


    
    
    Dim text As Variant
    Dim mc As MatchCollection, row As Long
    row = 2
    
    
    'Go through data row-wise and extract the email addresses
    
    For Each text In arr
        
        
        If regEx.Test(text) = True Then
            
            Set mc = regEx.Execute(text)
            Range("C" & row).Value = mc(0)
            
            
        End If
        
        row = row + 1
        
        
    Next text
   
   'Make things pretty (i.e. match the formatting of the other cells)
   
   With Range("C1")
    .Value = "Email"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    
   Range("C1:C51").Borders.LineStyle = xlContinuous
   Columns("C").AutoFit
   
   End With
    
End Sub
