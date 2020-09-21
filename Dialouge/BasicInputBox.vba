Option Explicit

Sub BasicInputbox()

    Dim yourName As String
    
    yourName = InputBox(Prompt:="What is your name", _
               Title:="Identify yourself")
              
    MsgBox ("Hello " & yourName & "!")
End Sub
