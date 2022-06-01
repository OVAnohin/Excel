Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Diagnostics
Imports System.IO
Imports System.Xml.Serialization
Imports System.Data.DataSetExtensions

Module Module1

    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Sub Main()


    End Sub

    Private Function GetNameSheet(fullFileName As String, sheetNumber As Integer) As String
        Dim oMissing As Object = System.Reflection.Missing.Value
        Dim excelApp As Excel.Application = New Excel.Application()
        Dim excelAppProcess As Process = GetExcelProcess(excelApp)
        excelApp.DisplayAlerts = False
        excelApp.FileValidationPivot = Excel.XlFileValidationPivotMode.xlFileValidationPivotRun
        Dim excelWb As Excel.Workbook = excelApp.Workbooks.Open(fullFileName)
        Dim excelWs As Excel.Worksheet = TryCast(excelWb.Worksheets(sheetNumber), Excel.Worksheet)

        Dim sheetName As String = excelWs.Name
        excelWb.Close(oMissing, oMissing, oMissing)
        excelApp.Quit()
        excelApp = Nothing
        excelAppProcess.Kill()

        ReleaseObject(excelApp)
        ReleaseObject(excelWb)
        ReleaseObject(excelWs)
        Return sheetName
    End Function

    Private Function GetExcelProcess(excelApp As Excel.Application) As Process
        Dim id As Integer
        GetWindowThreadProcessId(excelApp.Hwnd, id)
        Return Process.GetProcessById(id)
    End Function

    Private Sub ReleaseObject(comOj As Object)
        Try
            If comOj IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(comOj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comOj)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comOj)
            End If
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
