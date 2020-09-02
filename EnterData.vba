Sub EnterData()
'A easy Data entry form to embed in your excel spreadsheet
'can be assigned as a button or with a shortcut (I reccomend both)

'Error Handle

If Range("B2").Value = "" Then Exit Sub

' Enter new record
Range("F1048576").End(xlUp).Offset(1, 0).Select

' Copy record contents on sheet

ActiveCell.Value = Range("B2").Value
ActiveCell.Offset(0, 1).Value = Range("B3").Value
ActiveCell.Offset(0, 2).Value = Range("B4").Value
ActiveCell.Offset(0, 3).Value = Range("B5").Value

'Clear contents

Range("B2:B4").ClearContents

'Back to origin

Range("B2").Select



End Sub
