Attribute VB_Name = "ModuleConn"
Public CN1 As New ADODB.Connection
Public CN As New ADODB.Connection
Public Sub AttendConn()
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\main.mdb"
CN.Open
End Sub
Public Sub CloseConn()
CN.Close
Set CN = Nothing
End Sub
Public Sub ConnPass()
CN1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\pas.mdb"
CN1.Open
End Sub
Public Sub DissConnPass()
CN1.Close
Set CN1 = Nothing
End Sub
