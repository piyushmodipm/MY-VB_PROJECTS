Attribute VB_Name = "Module1"
Public Function getNewNo(tblName As String, colName As String) As Integer

Dim con As New Connection
Dim rs As New Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"

rs.Open "select " & colName & " from " & tblName & " order by " & colName, con, adOpenDynamic, adLockOptimistic


Dim n As Integer

If rs.BOF = True And rs.EOF = True Then
    n = 1
    getNewNo = n
    
Else
    rs.MoveLast
    n = Int(rs.Fields(0))
    getNewNo = n + 1
End If
rs.Close
con.Close
End Function

Public Function FillCombo(ByRef cmbBox As ComboBox, qry As String)
Dim con As New Connection
Dim rs As New Recordset
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"





Dim i As Integer
i = 0
cmbBox.Clear
While Not rs.EOF
    
    cmbBox.AddItem rs.Fields(1) & " : " & rs.Fields(0)
    cmbBox.ItemData(i) = rs.Fields(0)
    i = i + 1
    rs.MoveNext
Wend

End Function
