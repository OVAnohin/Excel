Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        Dim tableSummaryOfContracts As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\TableSummaryOfContracts.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableSummaryOfContracts.GetType())

        tableSummaryOfContracts = deSerializer.Deserialize(stream)
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

        'table TableBlockedUkr.xml
        Dim tableBlockedUkr As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlockedUkr.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlockedUkr.GetType())

        tableBlockedUkr = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableParkedUkr.xml
        Dim tableParkedUkr As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableParkedUkr.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableParkedUkr.GetType())

        tableParkedUkr = deSerializer.Deserialize(stream)
        stream.Close()

        '************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        If (tableSummaryOfContracts IsNot Nothing) Then
            If (tableSummaryOfContracts.Rows.Count > 0) Then
                tableSummaryOfContracts.Columns.Add("Запаркованные", Type.GetType("System.String"))
                tableSummaryOfContracts.Columns.Add("Блокированные", Type.GetType("System.String"))

                'по запаркованным
                For i As Integer = 0 To tableSummaryOfContracts.Rows.Count - 1
                    Dim row As DataRow = tableSummaryOfContracts.Rows(i)
                    If Not DBNull.Value.Equals(row("Поставщик")) AndAlso Not (row("БЕ") Like "UA*") Then
                        view = New DataView(tableParked)
                        Dim searchString As String = row("Поставщик")
                        filter = "[Код кредитора] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            row("Запаркованные") = tempTable.Rows(0)("Распределение мониторинга")
                        Else
                            row("Запаркованные") = "#Н/Д"
                        End If
                    End If
                Next
                'по Блокированные
                For i As Integer = 0 To tableSummaryOfContracts.Rows.Count - 1
                    Dim row As DataRow = tableSummaryOfContracts.Rows(i)
                    If Not DBNull.Value.Equals(row("Поставщик")) AndAlso Not (row("БЕ") Like "UA*") Then
                        view = New DataView(tableBlocked)
                        Dim searchString As String = row("Поставщик")
                        filter = "[Счет] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            row("Блокированные") = tempTable.Rows(0)("Распределение мониторинга")
                        Else
                            row("Блокированные") = "#Н/Д"
                        End If
                    End If
                Next

                'Ukr
                'по запаркованным Ukr
                For i As Integer = 0 To tableSummaryOfContracts.Rows.Count - 1
                    Dim row As DataRow = tableSummaryOfContracts.Rows(i)
                    If Not DBNull.Value.Equals(row("Поставщик")) AndAlso row("БЕ") Like "UA*" Then
                        view = New DataView(tableParkedUkr)
                        Dim searchString As String = row("Поставщик")
                        filter = "[Код кредитора] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            row("Запаркованные") = tempTable.Rows(0)("Код кредитора")
                        Else
                            row("Запаркованные") = "#Н/Д"
                        End If
                    End If
                Next
                'по Блокированные Ukr
                For i As Integer = 0 To tableSummaryOfContracts.Rows.Count - 1
                    Dim row As DataRow = tableSummaryOfContracts.Rows(i)
                    If Not DBNull.Value.Equals(row("Поставщик")) AndAlso row("БЕ") Like "UA*" Then
                        view = New DataView(tableBlockedUkr)
                        Dim searchString As String = row("Поставщик")
                        filter = "[Счет] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            row("Блокированные") = tempTable.Rows(0)("Код кредитора")
                        Else
                            row("Блокированные") = "#Н/Д"
                        End If
                    End If
                Next
            End If
        End If



        Console.ReadKey()

    End Sub

End Module
