Option Explicit

' A Subroutine for creating sheets for each month of the year
Sub createMonthSheets():
    Dim MonthNum As Integer
    
    Workbooks.Add
    
    For MonthNum = 1 To 12
    
    Worksheets.Add After:=Sheets(Sheets.Count)
    
    ActiveSheet.Name = MonthName(MonthNum, True)
    
    Next MonthNum
    

End Sub

'A Subroutine for creating sheets for up to this month

Sub createMonthSheets2()
   
   Dim MonthNum As Integer
    
    Workbooks.Add
    
    For MonthNum = 1 To 12
    
    Worksheets.Add After:=Sheets(Sheets.Count)
    
    ActiveSheet.Name = MonthName(MonthNum, True)
    
    ' Condition
    
    If MonthNum = Month(Date) Then Exit For
    
    Next MonthNum
    
End Sub
