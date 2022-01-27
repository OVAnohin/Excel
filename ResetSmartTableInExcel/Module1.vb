Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Module Module1

    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim fileName As String = "Блокированные.xlsx"
    Dim sheetName As String = "УП1"

    'out var
    Dim isComplete As Boolean = False
    Dim Message As String = ""

    Sub Main()

        Dim listOfWorkBooks As List(Of Workbook) = FindOpenedWorkBooks()

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = Reflection.Missing.Value
        Dim fullFileName As String = localFolder & "\" & fileName

        Try
            'xlApp = New Microsoft.Office.Interop.Excel.Application()
            'xlWorkBook = xlApp.Workbooks.Open(fullFileName)
            If listOfWorkBooks Is Nothing OrElse listOfWorkBooks.Count = 0 Then
                Exit Sub
            End If
            Console.WriteLine(listOfWorkBooks(0).Name)

            xlWorkBook = listOfWorkBooks(0)
            xlWorkBook.SaveAs("d:\Time\VB-Excel.xlsb", 50)
            xlWorkBook.Close(True, misValue, misValue)
            OnlyCloseExcelInstance()

            xlWorkSheet = CType(xlWorkBook.Sheets(sheetName), Excel.Worksheet)

            Dim selectedCell As Object
            Dim tableName As String

            selectedCell = xlWorkSheet.Range("A1")
            tableName = selectedCell.ListObject.Name
            Dim tbl As Object = xlWorkSheet.ListObjects(tableName)

            'Delete all table rows except first row
            If Not tbl.DataBodyRange Is Nothing Then
                If tbl.DataBodyRange.Rows.Count > 1 Then
                    tbl.AutoFilter.ShowAllData
                    tbl.DataBodyRange.Offset(1, 0).Resize(tbl.DataBodyRange.Rows.Count - 1, tbl.DataBodyRange.Columns.Count).Rows.Delete
                End If
                'Clear out data from first table row
                tbl.DataBodyRange.Rows(1).ClearContents
            End If

            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)

            isComplete = True

        Catch e As Exception
            isComplete = False
            Message = e.Message
        Finally
            xlApp = Nothing
            xlWorkBook = Nothing
            xlWorkSheet = Nothing
        End Try

    End Sub

    Private Function FindOpenedWorkBooks() As List(Of Workbook)
        Dim OpenedWorkBooks As New List(Of Workbook)()

        Dim ExcelInstances As Process() = Process.GetProcessesByName("EXCEL")
        If ExcelInstances.Length = 0 Then
            Return Nothing
        End If

        Dim ExcelInstance As Excel.Application = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        If ExcelInstance Is Nothing Then Return Nothing
        'Dim worksheets As Sheets = Nothing
        For Each WB As Workbook In ExcelInstance.Workbooks
            OpenedWorkBooks.Add(WB)
            'worksheets = WB.Worksheets
            'Console.WriteLine(WB.FullName)
            'For Each ws As Worksheet In worksheets
            '    Console.WriteLine(ws.Name)
            '    Marshal.ReleaseComObject(ws)
            'Next
        Next

        'Marshal.ReleaseComObject(worksheets)
        'worksheets = Nothing
        Marshal.FinalReleaseComObject(ExcelInstance)
        Marshal.CleanupUnusedObjectsInCurrentContext()
        ExcelInstance = Nothing
        Return OpenedWorkBooks
    End Function

    Private Sub OnlyCloseExcelInstance()
        Dim OpenedWorkBooks As New List(Of Workbook)()

        Dim ExcelInstances As Process() = Process.GetProcessesByName("EXCEL")
        If ExcelInstances.Length = 0 Then
            Exit Sub
        End If

        Dim ExcelInstance As Excel.Application = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        If ExcelInstance Is Nothing Then
            Exit Sub
        End If
        ExcelInstance.Quit()

        Marshal.FinalReleaseComObject(ExcelInstance)
        Marshal.CleanupUnusedObjectsInCurrentContext()
        ExcelInstance = Nothing
    End Sub

    Private Sub ReleaseObject(ByVal comOj As Object)
        Try
            Marshal.ReleaseComObject(comOj)
            Marshal.FinalReleaseComObject(comOj)
            comOj = Nothing
        Catch ex As Exception
            comOj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

End Module
