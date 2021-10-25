Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        Dim tableParkedBlocked As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\TableParkedBlocked.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableParkedBlocked.GetType())

        tableParkedBlocked = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableBlocked.xml
        Dim tableBlocked As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlocked.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlocked.GetType())

        tableBlocked = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableParked.xml
        Dim tableParked As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableParked.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableParked.GetType())

        tableParked = deSerializer.Deserialize(stream)
        stream.Close()

        '************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        If (tableParkedBlocked IsNot Nothing) Then
            If (tableParkedBlocked.Rows.Count > 0) Then
                tableParkedBlocked.Columns.Add("Запаркованные", Type.GetType("System.String"))
                tableParkedBlocked.Columns.Add("Блокированные", Type.GetType("System.String"))

                For i As Integer = 0 To tableParkedBlocked.Rows.Count - 1
                    Dim row As DataRow = tableParkedBlocked.Rows(i)
                    If Not DBNull.Value.Equals(row("№ документа")) AndAlso row("№ документа") <> Nothing AndAlso row("№ документа") <> "" Then
                        If Not (row("БЕ") Like "UA*") AndAlso row("Ст") = "@5C\Qоткрыт.@" Then
                            view = New DataView(tableBlocked)
                            Dim searchString As String = row("№ документа")
                            filter = "[№ документа] = '" & searchString & "'"
                            view.RowFilter = filter
                            tempTable = view.ToTable()
                            If tempTable.Rows.Count > 0 Then
                                row("Блокированные") = tempTable.Rows(0)("Распределение мониторинга")
                            Else
                                row("Блокированные") = "#Н/Д"
                            End If
                        End If
                    End If
                Next

                For i As Integer = 0 To tableParkedBlocked.Rows.Count - 1
                    Dim row As DataRow = tableParkedBlocked.Rows(i)
                    If Not DBNull.Value.Equals(row("№ документа")) AndAlso row("№ документа") <> Nothing AndAlso row("№ документа") <> "" Then
                        If Not (row("БЕ") Like "UA*") AndAlso row("Ст") = "@5D\QПредвРег@" Then
                            view = New DataView(tableBlocked)
                            Dim searchString As String = row("№ документа")
                            filter = "[№ документа] = '" & searchString & "'"
                            view.RowFilter = filter
                            tempTable = view.ToTable()
                            If tempTable.Rows.Count > 0 Then
                                row("Запаркованные") = tempTable.Rows(0)("Распределение мониторинга")
                            Else
                                row("Запаркованные") = "#Н/Д"
                            End If
                        End If
                    End If
                Next
            End If
        End If

    End Sub

End Module
