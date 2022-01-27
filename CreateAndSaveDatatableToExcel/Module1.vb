Imports System.IO
Imports System.Xml.Serialization
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Module Module1

    Sub Main()
        Dim _table As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\TableBlocked.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(_table.GetType())

        _table = deSerializer.Deserialize(stream)
        stream.Close()

        Dim localFolder As String = "D:\Time"
        Dim fileName As String = "test.xlsb"

        'begin
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Try

            Dim misValue As Object = System.Reflection.Missing.Value

            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)


            If (_table IsNot Nothing) Then
                If (_table.Rows.Count > 0) Then
                    Dim timeArray(_table.Rows.Count, _table.Columns.Count) As Object
                    Dim row As Integer, col As Integer

                    For row = 0 To _table.Rows.Count - 1
                        For col = 0 To _table.Columns.Count - 1
                            timeArray(row, col) = _table.Rows(row).Item(col)
                        Next
                    Next

                    col = 0
                    For Each column As DataColumn In _table.Columns
                        xlWorkSheet.Cells(1, col + 1) = column.ColumnName
                        col += 1
                    Next

                    xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(_table.Rows.Count + 1, _table.Columns.Count)).Value = timeArray
                End If
            End If
            xlWorkBook.SaveAs(localFolder & "\" & fileName, 50)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

        Catch e As Exception
            'Success = False
            'Message = e.Message
        Finally
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
        End Try

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module