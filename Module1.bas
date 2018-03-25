Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public n As ADODB.Recordset

Public Sub db()
Set con = New ADODB.Connection
Set n = New ADODB.Recordset
con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb"
End Sub
