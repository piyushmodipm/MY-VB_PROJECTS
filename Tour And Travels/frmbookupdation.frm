VERSION 5.00
Begin VB.Form frmbookupdation 
   Caption         =   "New Book Entry form"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmbookupdation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Dim con As New Connection
Dim rs As New Recordset

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient
rs.Open "select * from BookMaster", con, adOpenDynamic, adLockPessimistic



rs.Fields(0).Value = txtbookno.Text
Set DataGrid2.DataSource = rs
rs.Fields(1).Value = txttitle.Text
rs.Fields(2).Value = txtauthor.Text
rs.Fields(3).Value = txtprice.Text
rs.Fields(4).Value = txtedition.Text

rs.Update

MsgBox "Updation Successful"

End Sub

Private Sub cmdClose_Click()
Me.Hide
End Sub

