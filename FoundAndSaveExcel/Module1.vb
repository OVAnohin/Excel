Imports System.Runtime.InteropServices

Module Module1

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (hWnd1 As IntPtr, hWnd2 As IntPtr, lpsz1 As String, lpsz2 As String) As IntPtr
    Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (hWnd As IntPtr, dwId As Int32, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppvObject As Object) As Int32

    Private HandleMap As Dictionary(Of Integer, Object)
    Private InstanceMap As Dictionary(Of Object, Integer)
    Private CurrentInstance As Object

    Private Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
    Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0

    Sub Main()

        Dim Timeout As Integer = 30
        Dim handle As Integer
        Dim WorkbookName As String = "Microsoft Excel - work1"

        Dim obj = ExecWithTimeout(Timeout, "Open Named Instance", Function() OpenNamedObjectAA(WorkbookName))
        handle = GetHandle(obj)


    End Sub

    Private Function ExecWithTimeout(Of T)(timeout As Integer, name As String, operation As Func(Of T)) As T
        Dim ar = operation.BeginInvoke(Nothing, Nothing)
        If Not ar.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(timeout)) Then
            Throw New TimeoutException(name & " took more than " & timeout & " secs.")
        End If
        Return operation.EndInvoke(ar)
    End Function

    Private Sub ExecWithTimeout(timeout As Integer, name As String, operation As Action)
        Dim ar = operation.BeginInvoke(Nothing, Nothing)
        If Not ar.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(timeout)) Then
            Throw New TimeoutException(name & " took more than " & timeout & " secs.")
        End If
        operation.EndInvoke(ar)
    End Sub

    Private Function GetHandle(Instance As Object) As Integer

        If Instance Is Nothing Then
            Throw New ArgumentNullException("Tried to add an empty instance")
        End If

        ' Check if we already have this instance - if so, return it.
        If InstanceMap.ContainsKey(Instance) Then
            CurrentInstance = Instance
            Return InstanceMap(Instance)
        End If

        Dim key As Integer
        For key = 1 To Integer.MaxValue
            If Not HandleMap.ContainsKey(key) Then
                HandleMap.Add(key, Instance)
                InstanceMap.Add(Instance, key)
                CurrentInstance = Instance
                Return key
            End If
        Next key

        Return 0

    End Function

    Private Function OpenNamedObjectAA(workbookName As String) As Object
        Const OBJID_NATIVEOM = &HFFFFFFF0
        Dim IID_DISPATCH As New Guid("00020400-0000-0000-C000-000000000046")
        Dim workBook As Object = Nothing
        Do
            Dim XLhwnd As IntPtr = FindWindowEx(IntPtr.Zero, XLhwnd, "XLMAIN", Nothing)
            If IntPtr.Equals(XLhwnd, IntPtr.Zero) Then Exit Do
            Dim XLDESKhwnd As IntPtr = FindWindowEx(XLhwnd, IntPtr.Zero, "XLDESK", Nothing)
            Dim WBhwnd As IntPtr = FindWindowEx(XLDESKhwnd, IntPtr.Zero, "EXCEL7", Nothing)
            AccessibleObjectFromWindow(WBhwnd, OBJID_NATIVEOM, IID_DISPATCH, workBook)
            If workBook IsNot Nothing Then
                Dim application As Object = workBook.Application
                If application IsNot Nothing Then
                    Try
                        application.Windows(workbookName).Activate()
                        Return application
                    Catch ex As Exception
                        Continue Do
                    End Try
                End If
            End If
        Loop
        Throw New Exception("Excel with workbook name '" & workbookName & "' not found.")
    End Function



End Module
