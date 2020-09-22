Option Explicit

Sub DefaultCake()
    Dim CakeType As String
    
    CakeType = InputBox( _
    Prompt:="Which Cake do you want?", _
    Default:="Chocolate")
    
    'Some additional Dialouge Cancel is selected/ window is closed
    
    If CakeType = "" Then
    MsgBox "You didn't tell me what your favorite Cake!", vbCritical
    
    End If
    
End Sub

