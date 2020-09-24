Option Explicit

Sub ExcelUserForm()

    Dim YourName As String
    
    ' This is the Excel input box (not to be confused with a vba input box)
    ' Which is just `InputBox()` or `VBA.InputBox`
    YourName = Application.InputBox("Enter your name")
    
    MsgBox "Hello " & YourName & "!"
    
    
End Sub
