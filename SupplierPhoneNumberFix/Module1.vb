Imports System.Data.OleDb

Module Module1

    Sub Main()

        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\TableWork5.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Sheet1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim tableWork5 As DataTable
        oDataAdapter.Fill(oDataSet)
        tableWork5 = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\TableSupplierPhoneNumber.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        sheetName = "Sheet1"
        'sheetName = "Sheet1"
        oDataAdapter = New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        oDataSet = New DataSet
        Dim TableSupplierPhoneNumber As DataTable
        oDataAdapter.Fill(oDataSet)
        TableSupplierPhoneNumber = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        tempTable = TableSupplierPhoneNumber.Clone()

        'Filter = убираем пустые строки
        view = New DataView(TableSupplierPhoneNumber)
        filter = "[БЕ] <> ''"
        view.RowFilter = filter
        TableSupplierPhoneNumber = view.ToTable()

        For i As Integer = 0 To TableSupplierPhoneNumber.Rows.Count - 1
            Dim row As DataRow = TableSupplierPhoneNumber.Rows(i)
            If DBNull.Value.Equals(row("Ссылка на платеж")) Then
                Continue For
            End If
            Dim searchString As String = row("Ссылка на платеж")
            Dim searchChar = "/"
            Dim position = InStr(1, searchString, searchChar, 1)
            If position <> Nothing AndAlso position <> 0 Then
                row("Ссылка на платеж") = Mid(searchString, position + 1)
            End If
        Next

        For i As Integer = 0 To TableSupplierPhoneNumber.Rows.Count - 1
            Dim row As DataRow = TableSupplierPhoneNumber.Rows(i)
            If DBNull.Value.Equals(row("Номер договора")) Then
                Continue For
            End If
            Dim searchString As String
            If Not DBNull.Value.Equals(row("Ссылка на платеж")) Then
                searchString = row("Ссылка на платеж")
            End If

            Dim contractNumber As String = row("Номер договора")
            If contractNumber = Nothing Then
                Continue For
            End If
            If searchString = Nothing AndAlso contractNumber <> Nothing Then
                row("Ссылка на платеж") = row("Номер договора")
                Continue For
            End If
            If searchString <> Nothing AndAlso Left(searchString, 2) <> "46" Then
                row("Ссылка на платеж") = row("Номер договора")
            End If
        Next

        TableSupplierPhoneNumber.Columns.Add("AI", Type.GetType("System.String"))

        For i As Integer = 0 To TableSupplierPhoneNumber.Rows.Count - 1
            Dim row As DataRow = TableSupplierPhoneNumber.Rows(i)
            view = New DataView(tableWork5)

            If DBNull.Value.Equals(row("Ссылка на платеж")) Then
                row("AI") = "#Н/Д"
                Continue For
            End If

            If row("Ссылка на платеж").ToString() <> Nothing AndAlso row("Ссылка на платеж").ToString() <> "" Then
                filter = "[Номер Договора] = " & row("Ссылка на платеж")
                view.RowFilter = filter
                tempTable = view.ToTable()
                If tempTable.Rows.Count > 0 Then
                    row("AI") = tempTable.Rows(0)("Новое значение")
                Else
                    row("AI") = "#Н/Д"
                End If
            Else
                row("AI") = "#Н/Д"
            End If
        Next

        view = New DataView(TableSupplierPhoneNumber)
        filter = "[AI] <> '#Н/Д'"
        view.RowFilter = filter
        ShowTable(view.ToTable())
        'verificationTab = view.ToTable()

        'tableSupplierPhoneNumberOut = TableSupplierPhoneNumber

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
