Private Sub SqlFun()

'Using Early Binding
Dim connection As New ADODB.connection
' Gotta just copy this code down
connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName _
                 & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

Dim query As String

' Writing SQL!!!
query = "Select * from [Data$]"


Dim rs As New ADODB.Recordset

rs.Open query, connection

Sheet2.Range("A2").CopyFromRecordset rs

connection.Close

End Sub

Private Sub SqlFun2()

'Using Late Binding

Dim connection As Object

Set connection = CreateObject("ADODB.connection")

' Gotta just copy this code down
connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName _
                 & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

Dim query As String

' Writing SQL!!!
query = "Select * from [Data$]"


Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

rs.Open query, connection

Sheet2.Range("A2").CopyFromRecordset rs

connection.Close
End Sub
