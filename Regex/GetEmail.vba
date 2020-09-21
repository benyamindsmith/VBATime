Option Explicit

Function GetEmail(cell As String)
    
    '####################################
    'Make sure to first go to Tools > References...
    'and check off "Microsoft VBScript Regular Expressions 5.5
    '#####################################
    
    ' Create Regex object
    Dim regEx As New RegExp
    regEx.Global = True
    
    ' Define pattern
    regEx.Pattern = "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)"
            
    If regEx.Test(cell) = True Then
            
           GetEmail = regEx.Execute(cell)(0)
            
            
    End If
            
            
End Function
