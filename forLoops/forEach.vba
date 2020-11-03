Option Explicit

'For-Each Loop practice

' Get names of worksheets in current workbook

Sub workbookNames()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
