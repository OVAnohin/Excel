Module Module1

    Sub Main()
        Dim dataTable As DataTable = New DataTable("test")
        dataTable.Columns.Add("Name", Type.GetType("System.String"))
        dataTable.Columns.Add("ClassName", Type.GetType("System.String"))
        dataTable.Columns.Add("NativeWindowHandle", Type.GetType("System.Int32"))

        For i As Integer = 0 To 10
            Dim newRow As DataRow = dataTable.NewRow()
            newRow("Name") = "Name " & i
            newRow("ClassName") = "ClassName " & i
            newRow("NativeWindowHandle") = i
            dataTable.Rows.Add(newRow)
        Next
        For i As Integer = 0 To 10
            If i = 2 Or i = 3 Or i = 6 Or i = 8 Then
                Continue For
            End If
            Dim newRow As DataRow = dataTable.NewRow()
            newRow("Name") = "Name " & i
            newRow("ClassName") = "ClassName " & i
            newRow("NativeWindowHandle") = i
            dataTable.Rows.Add(newRow)
        Next

        ShowTable(dataTable)

        Dim view As New DataView(dataTable)
        view.Sort = "ClassName"
        'PrintTableOrView(view, "ClassName")
        'ShowTable(view)
        Dim newTable As DataTable = view.ToTable(True)
        ShowTable(newTable)
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
