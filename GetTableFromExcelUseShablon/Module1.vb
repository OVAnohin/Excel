Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Diagnostics
Imports System.IO
Imports System.Xml.Serialization
Imports System.Data.DataSetExtensions

Module Module1

    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Private resultTable As DataTable
    Sub Main()
        Dim tableShablon As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("c:\Temp\WorkDir\shablon.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableShablon.GetType())

        tableShablon = deSerializer.Deserialize(stream)
        stream.Close()

        Dim localFolder As String = "C:\Temp\WorkDir"
        Dim excelFileName As String = "NewСontracts.xlsb"
        '*********************** Begin
        Dim fullFileName As String = localFolder & "\" & excelFileName
        Dim sheetName As String = GetNameSheet(fullFileName, 1)

        Dim dataFromExcel As DataTable = New System.Data.DataTable()
        Dim connetionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";" + "Extended Properties='Excel 12.0 Xml;HDR=Yes;'"
        Dim sql As String = "Select * from[" & sheetName & "$]"

        Using oledbCnn = New OleDb.OleDbConnection(connetionString)
            Using oledbCmd = New OleDbCommand(sql, oledbCnn)
                Using oledbAdaper As OleDbDataAdapter = New OleDbDataAdapter(oledbCmd)
                    oledbAdaper.Fill(dataFromExcel)
                End Using
            End Using
            oledbCnn.Close()
        End Using

        For i As Integer = dataFromExcel.Rows.Count - 1 To 0 Step -1
            Dim row As DataRow = dataFromExcel.Rows(i)
            If row.Item(0) Is Nothing Then
                dataFromExcel.Rows.Remove(row)
            ElseIf row.Item(0).ToString = "" Then
                dataFromExcel.Rows.Remove(row)
            End If
        Next
        dataFromExcel.AcceptChanges()

        For i As Integer = 0 To dataFromExcel.Rows.Count - 1
            Dim row As DataRow = dataFromExcel.Rows(i)
            tableShablon.ImportRow(row)
        Next

        resultTable = tableShablon

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
