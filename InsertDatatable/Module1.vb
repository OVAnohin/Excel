Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module Module1

    Private oConnection As OleDbConnection

    Sub Main()
        Dim columnsDateToFix As DataTable = New DataTable()
        columnsDateToFix.Columns.Add("ColumnName", Type.GetType("System.String"))
        Dim newRow As DataRow = columnsDateToFix.NewRow()
        newRow("ColumnName") = "Дата"
        columnsDateToFix.Rows.Add(newRow)
        'newRow = columnsDateToFix.NewRow()

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

        Dim xlApp As Excel.Application = New Excel.Application()

        If xlApp Is Nothing Then
            Return
        End If

        'create Excel
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = Reflection.Missing.Value
        Dim localFolder As String = "d:\Work"
        Dim fileName As String = "result.xlsb"
        Dim fullFileName As String = localFolder & "\" & fileName
        Dim sheetName As String = "Коллекция"

        xlWorkBook = xlApp.Workbooks.Open(fullFileName)
        'xlWorkSheet = CType(xlWorkBook.Sheets(2), Excel.Worksheet)
        xlWorkSheet = CType(xlWorkBook.Sheets(sheetName), Excel.Worksheet)

        'fix time column
        Dim columnName As String = "Время"
        For i As Integer = 0 To dataTable.Rows.Count - 1
            Dim dRow As DataRow = dataTable.Rows(i)
            dRow(columnName) = Mid(dRow(columnName), 12)
        Next
        '/fix time column

        Dim timeArray(dataTable.Rows.Count, dataTable.Columns.Count) As Object
        Dim row As Integer, col As Integer

        For row = 0 To dataTable.Rows.Count - 1
            For col = 0 To dataTable.Columns.Count - 1
                timeArray(row, col) = dataTable.Rows(row).Item(col)
            Next
        Next

        col = 0
        For Each column As DataColumn In dataTable.Columns
            xlWorkSheet.Cells(1, col + 1) = column.ColumnName
            col += 1
        Next

        If columnsDateToFix.Rows.Count > 0 Then
            For Each column As DataRow In columnsDateToFix.Rows
                For col = 0 To dataTable.Columns.Count - 1
                    If xlWorkSheet.Cells(1, col + 1).Value = column("ColumnName") Then
                        xlWorkSheet.Cells(2, col + 1).NumberFormat = "dd/mm/yyyy"
                    End If
                Next
            Next
        End If



        xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(dataTable.Rows.Count + 1, dataTable.Columns.Count)).Value = timeArray

        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()

        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)

        'For i As Integer = 1 To dataTable.Rows.Count
        '    'xlWorkSheet.Range("a" & (dataTable.Rows.Count + 2), "b" & i).NumberFormat = "mm/dd/yyyy"
        '    Console.WriteLine("a" & (dataTable.Rows.Count + 2) & " " & "b" & i)
        'Next

        'Console.ReadKey()

    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub DataTableToExcel(dataTable As DataTable, ws As Excel.Worksheet, sheetName As String)
        Dim arr(dataTable.Rows.Count, dataTable.Columns.Count) As Object
        Dim r As Integer, c As Integer

        For r = 0 To dataTable.Rows.Count - 1
            For c = 0 To dataTable.Columns.Count - 1
                arr(r, c) = dataTable.Rows(r).Item(c)
            Next
        Next

        ws.Name = sheetName
        c = 0
        For Each column As DataColumn In dataTable.Columns
            ws.Cells(1, c + 1) = column.ColumnName
            c += 1
        Next

        ws.Range(ws.Cells(2, 1), ws.Cells(dataTable.Rows.Count, dataTable.Columns.Count)).Value = arr
    End Sub

    Private Sub PrintTableOrView(ByVal dv As DataView)
        Dim sw As System.IO.StringWriter
        Dim output As String
        Dim table As DataTable = dv.Table

        For Each rowView As DataRowView In dv
            sw = New System.IO.StringWriter

            For Each col As DataColumn In table.Columns
                sw.Write(rowView(col.ColumnName).ToString() & ", ")
            Next
            output = sw.ToString
            If output.Length > 2 Then
                output = output.Substring(0, output.Length - 2)
            End If
            Console.WriteLine(output)
        Next

        Console.WriteLine(New String("*", 20))
    End Sub

    Private Sub PrintTableOrView(ByVal table As DataTable)
        Dim sw As System.IO.StringWriter
        Dim output As String

        For Each row As DataRow In table.Rows
            sw = New System.IO.StringWriter
            For Each col As DataColumn In table.Columns
                sw.Write(row(col).ToString() & ", ")
            Next
            output = sw.ToString
            If output.Length > 2 Then
                output = output.Substring(0, output.Length - 2)
            End If
            Console.WriteLine(output)
        Next

        Console.WriteLine(New String("*", 20))
    End Sub

End Module
