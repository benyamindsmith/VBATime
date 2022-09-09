Option Explicit

'
' A lookup function for attribute scoring
'
'
Function ILOOKUP(lookup_value As Range, _
                 lookup_range As Range, _
                 return_range As Range _
                 ) As String
Dim i As Integer
Dim lookup_ind As Integer
Dim lookup_val As Double
For i = 1 To lookup_range.Count
    lookup_ind = i
    lookup_val = lookup_range.Rows.Item(i)
    
If lookup_val <= lookup_value.Rows.Item(1) Then Exit For
Next

ILOOKUP = return_range.Rows.Item(lookup_ind)
End Function
