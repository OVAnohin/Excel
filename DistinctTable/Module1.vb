Imports System.Data.OleDb

Module Module1

    Sub Main()
        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\Work3.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Sheet1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim dataTable As DataTable
        oDataAdapter.Fill(oDataSet)
        dataTable = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        Dim view As New DataView(dataTable)
        view.Sort = "Номер Договора, Дата DESC, Время DESC"
        Dim newTable As DataTable = view.ToTable()

        Console.WriteLine()

        For i As Integer = 0 To newTable.Rows.Count - 2
            Dim currentRow As DataRow = newTable.Rows(i)
            Dim rowPlusOne As DataRow = newTable.Rows(i + 1)
            If currentRow(newTable.Columns("Номер Договора")) = rowPlusOne(newTable.Columns("Номер Договора")) Then
                rowPlusOne(newTable.Columns("Новое значение")) = currentRow(newTable.Columns("Новое значение"))
                rowPlusOne(newTable.Columns("Старое значение")) = currentRow(newTable.Columns("Старое значение"))
                rowPlusOne(newTable.Columns("Время")) = currentRow(newTable.Columns("Время"))
            End If
        Next

        view = New DataView(newTable)
        Dim result As DataTable
        result = view.ToTable(True)

        ShowTable(result)

        Console.ReadKey()
    End Sub

    Private Sub ShowTable(table As DataTable)
        For i = 0 To table.Rows.Count - 1
            For j = 0 To table.Columns.Count - 1
                Dim row As DataRow = table.Rows(i)
                Dim column As DataColumn = table.Columns(j)
                Console.Write(row(column) & " ")
            Next
            Console.WriteLine()
        Next
        Console.WriteLine(New String("*", 20))
    End Sub

End Module
