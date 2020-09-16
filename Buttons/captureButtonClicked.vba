Option Explicit

Sub captureButtonClicked()
    Dim ButtonClicked As VbMsgBoxResult
    
    ButtonClicked = MsgBox( _
    Prompt:="With me so far", _
    Buttons:=vbYesNo + vbQuestion)
    
    If ButtonClicked = vbYes Then
        MsgBox "Carry on!"
    Else
        MsgBox "Try this again"
    End If
    
End Sub
