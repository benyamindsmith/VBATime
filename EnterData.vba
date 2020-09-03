Sub EnterData()
'A easy Data entry form to embed in your excel spreadsheet
'can be assigned as a button or with a shortcut (I reccomend both)

'Error Handling

If Range("B2").Value = "" Then
    Range("B2").Interior.Color = rgbPink
    Range("B2").Select
    
    With Range("A5")
    .Value = "Enter Name"
    .Font.Color = rgbRed
    End With
    
    Exit Sub
End If

' Enter new record
Range("F1048576").End(xlUp).Offset(1, 0).Select

' Copy record contents on sheet

ActiveCell.Value = Range("B2").Value
ActiveCell.Offset(0, 1).Value = Range("B3").Value
ActiveCell.Offset(0, 2).Value = Range("B4").Value
ActiveCell.Offset(0, 3).Value = Range("B5").Value

'Clear contents

Range("B2:B4").ClearContents
Range("A5").Clear
Range("B2").Interior.Color = xlNone

'Back to origin

Range("B2").Select



End Sub
