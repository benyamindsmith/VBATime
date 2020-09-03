' A random score generator that lists the winner
Sub PredictScore()
        Range("A3").Value = WorksheetFunction.RandBetween(0, 5)
        Range("B3").Value = WorksheetFunction.RandBetween(0, 5)
        
        'List winner
        With Range("B5")
        .Value = IIf( _
            Range("A3").Value > Range("B3").Value, _
            Range("A2").Value, _
            IIf(Range("B3").Value > Range("A3").Value, _
            Range("A2").Value, _
            "Tie"))
        End With
End Sub
