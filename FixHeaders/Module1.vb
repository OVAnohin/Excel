Imports Microsoft.Office.Interop

Module Module1

    Sub Main()
        Dim fileName As String = "ContractualCost.xlsb"
        Dim localFolder As String = "d:\Work"
        Dim fullFileName As String = localFolder & "\" & fileName
        Dim row As Integer = 1
        Dim column As Integer = 5
        Dim value As String = "Пользоват"

        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlWorkBook = xlApp.Workbooks.Open(fullFileName)
        xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)

        xlWorkSheet.Cells(row, column).NumberFormat = "@"
        xlWorkSheet.Cells(row, column) = value
        xlWorkBook.Close(True)
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
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

End Module
