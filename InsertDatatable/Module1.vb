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
        'xlWorkSheet.Name = "Sheet1"
        'xlWorkSheet.Cells(1, 1) = "Sheet 1 content"
        DataTableToExcel(dataTable, xlWorkSheet, xlWorkSheet.Name)

        'xlWorkBook.SaveAs(fullFileName, 50)
        'xlWorkBook.Close(True, misValue, misValue)
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()

        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)

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
