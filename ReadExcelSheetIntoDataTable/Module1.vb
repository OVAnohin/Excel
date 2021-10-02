Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module Module1

    Private oConnection As OleDbConnection

    Sub Main()
        'datatable
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Work\ContractualCost.xlsb;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim oDataAdapter As New OleDbDataAdapter("Select * from[Лист1$]", oConnection)
        Dim oDataSet As New DataSet
        Dim dataTable As DataTable
        oDataAdapter.Fill(oDataSet)
        dataTable = oDataSet.Tables(0)
        oConnection.Close()

        ShowTable(dataTable)

        Dim fileName As String = "ContractualCost1.xlsb"
        Dim localFolder As String = "d:\Work"
        Dim fullFileName As String = localFolder & "\" & fileName
        Dim columnName As String = "Пользоват#"
        Dim columnNumber As Integer = 5

        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlWorkBook = xlApp.Workbooks.Open(fullFileName)
        xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)
        Dim excelRange As Excel.Range = xlWorkSheet.UsedRange
        Dim rows As Integer = excelRange.Rows.Count
        Dim cols As Integer = excelRange.Columns.Count

        For i As Integer = 2 To rows
            dataTable.Rows(i - 2)(columnName) = excelRange.Cells(i, columnNumber).Value2
        Next

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        ShowTable(dataTable)

        Console.ReadKey()

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
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
