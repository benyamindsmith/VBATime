Options Explicit

Sub getUserName()
    Dim myMsg As String
    
    myMsg = "Your username is: " & Environ("UserName")
    
    MsgBox myMsg
    
End Sub
