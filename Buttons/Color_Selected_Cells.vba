Sub Color_Selected_Cells()
    ' Color Selected Cells With a Random Color
    
    Selection.Interior.Color = WorksheetFunction.RandBetween(0, 16777215)

End Sub
