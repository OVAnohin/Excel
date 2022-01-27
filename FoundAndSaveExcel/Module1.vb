Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Windows.Automation

Module Module1

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function SendMessageW Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim fileName As String = "Блокированные11.xlsb"
    Dim sheetName As String = "УП1"

    Private Const WM_COMMAND = &H111
    Private Const BM_CLICK As Integer = &HF5

    'out var
    Dim isComplete As Boolean = False


    Sub Main()
        Dim isExit As Boolean = False
        Dim timeout As DateTime = DateTime.Now.AddSeconds(10)
        Dim excelProcesses As Process()
        Dim misValue As Object = Reflection.Missing.Value
        Dim exceptionMessage As String = ""

        While (isExit = False)
            excelProcesses = Process.GetProcessesByName("EXCEL")
            If excelProcesses.Length = 0 Then
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти процесс Excel.")
                End If
            Else
                ReleaseObject(excelProcesses)
                excelProcesses = Nothing
                isExit = True
            End If
        End While

        Dim xlApplication As Excel.Application
        isExit = False
        timeout = DateTime.Now.AddSeconds(30)
        While (isExit = False)
            Try
                Thread.Sleep(500)
                For Each app As Process In Process.GetProcessesByName("EXCEL")
                    Dim ptrWindow As IntPtr = FindWindow(Nothing, app.MainWindowTitle)
                    If ptrWindow <> IntPtr.Zero Then
                        ShowHideWindow(ptrWindow)
                        ''BringWindowToTop(hWnd)
                    End If
                Next
                Console.WriteLine("Before xlApp")
                'xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
                xlApplication = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                Console.WriteLine("After xlApp")
                If xlApplication Is Nothing Then
                    Continue While
                End If
                For Each xlWorkBook As Workbook In xlApplication.Workbooks
                    xlWorkBook.SaveAs(localFolder & "\" & fileName, 50)
                    xlWorkBook.Close(True)
                    ReleaseObject(xlWorkBook)
                    xlWorkBook = Nothing
                Next
                xlApplication.Quit()
                ReleaseObject(xlApplication)
                isExit = True
            Catch ex As Exception
                TryLaunchIE()
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти Excel.Application. " & exceptionMessage)
                End If
            Finally
                xlApplication = Nothing
            End Try
        End While

        Dim proc As Process
        For Each proc In Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next

        isComplete = True

    End Sub

    Private Sub ShowHideWindow(hWindow As IntPtr)
        Dim autoElement As AutomationElement = AutomationElement.FromHandle(hWindow)
        Dim elementCollectionAll As AutomationElementCollection = autoElement.FindAll(TreeScope.Subtree, Condition.TrueCondition)
        SetFocusOnWindow(elementCollectionAll)
        Dim ptrWindow As Integer = CType(hWindow, Integer)
        ShowWindow(ptrWindow, 0)
        Thread.Sleep(300)
        ShowWindow(ptrWindow, 9)
        Thread.Sleep(300)
        ShowWindow(ptrWindow, 3)
        Thread.Sleep(300)
        SendMessageW(ptrWindow, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
        Thread.Sleep(300)
    End Sub

    Private Sub TryLaunchIE()
        Dim ie As Process = Process.Start("iexplore.exe", "localhost")
        'close the website
        Thread.Sleep(2000)
        Try
            Dim ieMainWindow As AutomationElement = AutomationElement.FromHandle(ie.MainWindowHandle)
            Dim elementCollectionAll As AutomationElementCollection = ieMainWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            SetFocusOnWindow(elementCollectionAll)

            Thread.Sleep(200)
            Dim ieProc As Process
            For Each ieProc In Process.GetProcessesByName("iexplore")
                ieProc.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Function SetFocusOnWindow(elementCollectionAll As AutomationElementCollection) As Boolean

        For Each autoElement As AutomationElement In elementCollectionAll
            autoElement.SetFocus()
            Return True
        Next

        Return False
    End Function

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
