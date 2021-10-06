Imports System.Data.OleDb

Module Module1

    Sub Main()
        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\EndExecutionTime.xlsb;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Лист1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim tableEndExecutionTime As DataTable
        oDataAdapter.Fill(oDataSet)
        tableEndExecutionTime = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        tableEndExecutionTime.Columns.Add("Сцепить", Type.GetType("System.String"))

        For i As Integer = 0 To tableEndExecutionTime.Rows.Count - 1
            Dim row As DataRow = tableEndExecutionTime.Rows(i)
            row("Сцепить") = row("Блк") & "_" & row("Текст заголовка документа")
        Next

        Dim tableParkedBlocked As DataTable
        tableParkedBlocked = tableEndExecutionTime.Clone()
        Dim tableTemp2 As DataTable
        tableTemp2 = tableEndExecutionTime.Clone()
        Dim tableTemp3 As DataTable
        tableTemp3 = tableEndExecutionTime.Clone()

        'Filter = "@5C\Qоткрыт.@"
        view = New DataView(tableEndExecutionTime)
        filter = "[Ст] = '@5C\Qоткрыт.@'"
        view.RowFilter = filter
        Dim tableOpenPosition As DataTable = view.ToTable()

        'Filter = "@5D\QПредвРег@"
        view = New DataView(tableEndExecutionTime)
        filter = "[Ст] = '@5D\QПредвРег@'"
        view.RowFilter = filter
        Dim tablePreRegistration As DataTable = view.ToTable()

        view = New DataView(tableOpenPosition)
        filter = "[Сцепить] Like 'Y%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp2.ImportRow(tempTable.Rows(i))
        Next

        'затем по значению «W*,03,*» добавляем таблицу tableTemp3 !!!!
        view = New DataView(tableOpenPosition)
        filter = "[Сцепить] Like 'W%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp3.ImportRow(tempTable.Rows(i))
        Next

        view = New DataView(tableTemp3)
        filter = "[Сцепить] Like '%,03,%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp2.ImportRow(tempTable.Rows(i))
        Next

        view = New DataView(tableTemp3)
        filter = "[Сцепить] Like '%,03'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp2.ImportRow(tempTable.Rows(i))
        Next

        Console.WriteLine(tableTemp2.Rows.Count)

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
