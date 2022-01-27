Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Module Module1

    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim fileName As String = "Блокированные11.xlsb"
    Dim sheetName As String = "УП1"

    'out var
    Dim isComplete As Boolean = False
    Dim exceptionMessage As String = ""

    Sub Main()

        Dim listOfWorkBooks As List(Of Workbook) = FindOpenedWorkBooks()
        Dim xlWorkBook As Excel.Workbook
        Dim misValue As Object = Reflection.Missing.Value

        Try
            Console.WriteLine(listOfWorkBooks.Count)
            If listOfWorkBooks Is Nothing OrElse listOfWorkBooks.Count = 0 Then
                Exit Sub
            End If
            'Console.WriteLine(listOfWorkBooks(0).Name)
            xlWorkBook = listOfWorkBooks(0)
            xlWorkBook.SaveAs(localFolder & "\" & fileName, 50)
            xlWorkBook.Close(True, misValue, misValue)
            OnlyCloseExcelInstance()

            ReleaseObject(xlWorkBook)
            isComplete = True

        Catch e As Exception
            isComplete = False
            exceptionMessage = e.Message
        Finally
            'xlApp = Nothing
            xlWorkBook = Nothing
            'xlWorkSheet = Nothing
        End Try

    End Sub

    Private Function FindOpenedWorkBooks() As List(Of Workbook)
        Dim openedWorkBooks As New List(Of Workbook)()

        'Dim ExcelInstances As Process() = Process.GetProcessesByName("EXCEL")
        'If ExcelInstances.Length = 0 Then
        '	Throw New Exception("Не могу найти процесс Excel.")
        'End If

        Dim isExit As Boolean = False
        Dim timeout As DateTime = DateTime.Now.AddSeconds(10)
        Dim listExcelInstances As Process()

        While (isExit = False)
            listExcelInstances = Process.GetProcessesByName("EXCEL")
            If listExcelInstances.Length = 0 Then
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти процесс Excel.")
                End If
            Else
                ReleaseObject(listExcelInstances)
                listExcelInstances = Nothing
                isExit = True
            End If
        End While

        isExit = False
        timeout = DateTime.Now.AddSeconds(10)

        While (isExit = False)
            Try
                Dim xlAppTime As Excel.Application = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                ReleaseObject(xlAppTime)
                Marshal.ReleaseComObject(xlAppTime)
                xlAppTime = Nothing
                isExit = True
            Catch ex As Exception
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти Excel.Application.")
                End If
            End Try
        End While

        'Dim worksheets As Sheets = Nothing
        Dim xlApp As Excel.Application = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        For Each WB As Workbook In xlApp.Workbooks
            openedWorkBooks.Add(WB)
            'worksheets = WB.Worksheets
            'For Each ws As Worksheet In worksheets
            '    Marshal.ReleaseComObject(ws)
            'Next
        Next

        'ReleaseObject(worksheets)
        ReleaseObject(xlApp)
        Marshal.CleanupUnusedObjectsInCurrentContext()
        'worksheets = Nothing
        xlApp = Nothing

        Return openedWorkBooks
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
