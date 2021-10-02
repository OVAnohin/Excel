Imports System.Data.OleDb

Module Module1
    Private oConnection As OleDbConnection

    Private Class Work1
        Public Property Work1 As List(Of Work1Columns)
    End Class

    Private Class Work1Columns
        Public Property Creditor As String
        Public Property NameCreditor As String
        Public Property BE As String
        Public Property ContractNumber As String
        Public Property NewValue As String
        Public Property dDate As Date
        Public Property tTime As String
    End Class

    Sub Main()

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
        oConnection.Close()

        'ShowTable(dataTable)
        Dim view As New DataView(dataTable)
        view.Sort = "Номер Договора, Дата, Время DESC"
        Dim newTable As DataTable = view.ToTable()
        'ShowTable(newTable)
        For i = 0 To newTable.Rows.Count - 2
            Dim currentRow As DataRow = newTable.Rows(i)
            Dim rowPlusOne As DataRow = newTable.Rows(i + 1)
            If currentRow(newTable.Columns("Номер Договора")) = rowPlusOne(newTable.Columns("Номер Договора")) Then
                rowPlusOne(newTable.Columns("Новое значение")) = currentRow(newTable.Columns("Новое значение"))
                rowPlusOne(newTable.Columns("Старое значение")) = currentRow(newTable.Columns("Старое значение"))
                rowPlusOne(newTable.Columns("Время")) = currentRow(newTable.Columns("Время"))
            End If
        Next
        'ShowTable(newTable)
        view = New DataView(newTable)
        newTable = view.ToTable(True)
        'ShowTable(newTable)
        'Dim Amounts As New DataTable 'Your code to load actual DataTable here
        'Dim amountGrpByDates = From row In newTable
        '                       Group row By dateGroup = New With {
        '                                            Key .Creditor = row.Field(Of String)("Кредитор"),
        '                                            Key .NameCreditor = row.Field(Of String)("Наименование кредитора"),
        '                                            Key .BE = row.Field(Of String)("БЕ"),
        '                                            Key .ContractNumber = row.Field(Of String)("Номер Договора"),
        '                                            Key .NewValue = row.Field(Of String)("Новое значение"),
        '                                            Key .dDate = row.Field(Of String)("Дата"),
        '                                            Key .tTime = row.Field(Of String)("Время")
        '                                       } Into Group
        '                       Select New With {
        '                          Key .Dates = dateGroup,
        '                              .SumAmount = Group.Sum(Function(x) x.Field(Of String)("Номер Договора"))}

        'ShowTable(amountGrpByDates)

        'Dim dt As DataTable = dataTable.AsEnumerable().GroupBy(r >= New {Col1 = r["Col1"], Col2 = r["Col2"]}).Select(g >= g.OrderBy(r >= r["PK"]).First()).CopyToDataTable()

        'Dim view As New DataView(dataTable)
        ''view.Sort = "Номер Договора, Дата, Время"
        '''PrintTableOrView(view, "ClassName")
        '''ShowTable(view)
        ''Dim newTable As DataTable = view.ToTable(True, {"Кредитор", "Номер Договора", "Дата"})
        'ShowTable(newTable)

        'Dim work1 As New Work1
        'work1.Work1 = ConvertTableToList(dataTable)

        'For i = 0 To 10
        '    Console.WriteLine(i & " - " & work1.Work1.Item(i).ContractNumber)
        'Next

        Console.ReadKey()

    End Sub

    Private Function ConvertTableToList(dt As DataTable) As List(Of Work1Columns)

        Dim rowsWork1 As New List(Of Work1Columns)
        For Each rw As DataRow In dt.Rows
            Dim payment As New Work1Columns With
            {
                .Creditor = rw.Item("Кредитор"),
                .NameCreditor = rw.Item("Наименование кредитора"),
                .BE = rw.Item("БЕ"),
                .ContractNumber = rw.Item("Номер Договора"),
                .NewValue = rw.Item("Новое значение"),
                .dDate = rw.Item("Дата"),
                .tTime = rw.Item("Время")
            }
            rowsWork1.Add(payment)
        Next

        Return rowsWork1

    End Function

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

    Private Sub PrintTableOrView(ByVal dv As DataView, ByVal label As String)
        Dim sw As System.IO.StringWriter
        Dim output As String
        Dim table As DataTable = dv.Table

        Console.WriteLine(label)

        ' Loop through each row in the view.
        For Each rowView As DataRowView In dv
            sw = New System.IO.StringWriter

            ' Loop through each column.
            For Each col As DataColumn In table.Columns
                ' Output the value of each column's data.
                sw.Write(rowView(col.ColumnName).ToString() & ", ")
            Next
            output = sw.ToString
            ' Trim off the trailing ", ", so the output looks correct.
            If output.Length > 2 Then
                output = output.Substring(0, output.Length - 2)
            End If
            ' Display the row in the console window.
            Console.WriteLine(output)
        Next
        Console.WriteLine()
    End Sub

    Private Sub PrintTableOrView(ByVal table As DataTable, ByVal label As String)
        Dim sw As System.IO.StringWriter
        Dim output As String

        Console.WriteLine(label)

        ' Loop through each row in the table.
        For Each row As DataRow In table.Rows
            sw = New System.IO.StringWriter
            ' Loop through each column.
            For Each col As DataColumn In table.Columns
                ' Output the value of each column's data.
                sw.Write(row(col).ToString() & ", ")
            Next
            output = sw.ToString
            ' Trim off the trailing ", ", so the output looks correct.
            If output.Length > 2 Then
                output = output.Substring(0, output.Length - 2)
            End If
            ' Display the row in the console window.
            Console.WriteLine(output)
        Next
        Console.WriteLine()
    End Sub

End Module
