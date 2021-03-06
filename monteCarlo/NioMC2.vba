Option Explicit

Sub displayResults()
 On Error Resume Next
 Worksheets("Result").Delete
 On Error GoTo 0
 Worksheets.Copy After:=Worksheets("NIO")
 ActiveSheet.Name = "Result"
 
 'Preliminaries:
 'Make sure data is sorted from newest to oldest
 'Column A- Your Date Values
 'Column B- Your Open values
 
 'Name column with computations
 
 Range("C1").Value = "Change"
 
 

 Dim rws As Range
 Dim changeVals As Range
 Dim rowLength As Integer
 
 Set rws = Range("B2:B557")
 Set changeVals = Range("C2:C557")
 rowLength = rws.Count
 
 
 
 ' Calculate change from today to yesterday
 On Error Resume Next
 Dim i As Integer
 For i = 1 To rowLength
 With Range("C" & i + 1)
 .Value = WorksheetFunction.Ln(Range("B" & i + 1).Value / Range("B" & i + 2).Value)
 End With
 Next i
 On Error GoTo 0
 
 'Do monte Carlo Simulation
 
 'Define number of simulations
 
 Dim simNumber As Integer
 Dim t As Integer
 simNumber = 100
 'iteration
 For t = 1 To simNumber
 'the actual monte carlo
 Range("D1").Value = "MC Estimate"
 Range("D2").Value = Range("B2").Value
 
 Dim j As Integer
 For j = 1 To rowLength
 
 'Do loop error handles
 Do
 On Error Resume Next
 With Range("D" & j + 2)
 
  .Value = Range("D" & j + 1) * _
    Exp( _
            WorksheetFunction.Small(changeVals, WorksheetFunction.RandBetween(1, rowLength)))
 End With
 Loop Until (Err.Number = 0)
 
 
 Next j
 If t < 100 Then
 Range("D:D").EntireColumn.Insert
 End If
 Next t


End Sub


