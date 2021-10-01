Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module Module1

    Private oConnection As OleDbConnection

    Sub Main()

        'datatable
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\work1.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim oDataAdapter As New OleDbDataAdapter("Select * from[Sheet1$]", oConnection)
        Dim oDataSet As New DataSet
        Dim dataTable As DataTable
        oDataAdapter.Fill(oDataSet)
        dataTable = oDataSet.Tables(0)
        'PrintTableOrView(dataTable)

    End Sub

End Module
