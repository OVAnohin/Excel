Imports System.Data.OleDb

Module Module1

    Sub Main()
        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\Work5.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Sheet1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim work5 As DataTable
        oDataAdapter.Fill(oDataSet)
        work5 = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        work5.Columns.Add("YesOrNo", Type.GetType("System.String"))

        For i As Integer = 0 To work5.Rows.Count - 1
            Dim row As DataRow = work5.Rows(i)
            'тут пусто
            If DBNull.Value.Equals(row("Новое значение")) AndAlso DBNull.Value.Equals(row("Старое значение")) Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'одно пустое, второе есть, на проверку
            If Not DBNull.Value.Equals(row("Новое значение")) AndAlso DBNull.Value.Equals(row("Старое значение")) Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'одно пустое, второе есть, на проверку (reverse)
            If DBNull.Value.Equals(row("Новое значение")) AndAlso Not DBNull.Value.Equals(row("Старое значение")) Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'одно пустое, второе есть, на проверку
            If row("Новое значение") = "" AndAlso row("Старое значение") <> "" Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'одно пустое, второе есть, на проверку (reverse)
            If row("Новое значение") <> "" AndAlso row("Старое значение") = "" Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'анализ на то что просто равны
            If row("Новое значение") = row("Старое значение") Then
                row("YesOrNo") = "No"
                Continue For
            End If

            Dim arrNewValue As String()
            Dim arrOldValue As String()
            arrNewValue = Split(row("Новое значение"), ",")
            arrOldValue = Split(row("Старое значение"), ",")

            'первые значения массивов равны = не нужно проверять
            If arrNewValue(0) = arrOldValue(0) Then
                row("YesOrNo") = "No"
                Continue For
            End If
            'елси первое значение есть в списке второго
            If Array.IndexOf(arrOldValue, arrNewValue(0)) <> -1 Then
                row("YesOrNo") = "No"
                Continue For
            End If
            row("YesOrNo") = "Yes"
        Next

        Dim view As New DataView(work5)
        view.RowFilter = "[YesOrNo] = 'Yes'"
        Dim newTable As DataTable = view.ToTable()

        newTable.Columns.Remove("YesOrNo")

        ShowTable(newTable)
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
