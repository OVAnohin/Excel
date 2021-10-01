Imports Microsoft.Office.Interop

Module Module1

    Sub Main()

        Dim xlApp As Excel.Application = New Excel.Application()

        If xlApp Is Nothing Then
            Console.WriteLine("Excel is not properly installed!!")
            Return
        End If

        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = Reflection.Missing.Value

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        'Console.WriteLine(xlWorkBook.Sheets(0).Name)
        xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)
        'xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.Cells(1, 1) = "Sheet 1 content"

        xlWorkBook.SaveAs("d:\Time\VB-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)

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

End Module
