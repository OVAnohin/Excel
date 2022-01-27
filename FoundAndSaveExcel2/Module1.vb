Module Module1

    Private mLastException As Exception = Nothing
    Private InstanceMap As Dictionary(Of Object, Integer)
    Private CurrentInstance As Object
    Private HandleMap As Dictionary(Of Integer, Object)

    Sub Main()
        Dim timeout As Double = 3000
        Dim obj As Object = Nothing
        Dim handle As Double
        Try
            obj = ExecWithTimeout(timeout, "Get Object", Function() GetObject(, "Excel.Application"))
        Catch ex As TimeoutException
            mLastException = ex
            Throw
        End Try
        If obj Is Nothing Then Throw New Exception("Could not Get Object")

        ' GetObject may return an unusable wrapper (possibly if instance is shutting
        ' down) which results in a "COM target does not implement IDispatch" exception
        ' when accessing members of the object. If reading the EnableEvents property 
        ' results in an exception, the recover stage will run and a new instance will 
        ' be created instead.
        Dim tempEnableEvents = obj.EnableEvents

        handle = GetHandle(obj)

    End Sub

    Private Function ExecWithTimeout(Of T)(timeout As Integer, name As String, operation As Func(Of T)) As T
        Dim ar = operation.BeginInvoke(Nothing, Nothing)
        If Not ar.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(timeout)) Then
            Throw New TimeoutException(name & " took more than " & timeout & " secs.")
        End If
        Return operation.EndInvoke(ar)
    End Function

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

End Module
