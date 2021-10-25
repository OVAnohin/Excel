Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        'table TableSummaryOfPrice.xml
        Dim tableSummaryOfPrice As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\TableSummaryOfPrice.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableSummaryOfPrice.GetType())

        tableSummaryOfPrice = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableBlocked2.xml
        Dim tableBlocked2 As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlocked2.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlocked2.GetType())

        tableBlocked2 = deSerializer.Deserialize(stream)
        stream.Close()
        tableBlocked2.Clear()

        'table TableBlockedUkr2.xml
        Dim tableBlockedUkr2 As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlockedUkr2.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlockedUkr2.GetType())

        tableBlockedUkr2 = deSerializer.Deserialize(stream)
        stream.Close()
        tableBlockedUkr2.Clear()

        'table TableBlocked.xml
        Dim tableBlocked As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlocked.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlocked.GetType())

        tableBlocked = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableBlockedUkr.xml
        Dim tableBlockedUkr As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlockedUkr.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlockedUkr.GetType())

        tableBlockedUkr = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableExceptionsForCreditors.xml
        Dim tableExceptionsForCreditors As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableExceptionsForCreditors.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableExceptionsForCreditors.GetType())

        tableExceptionsForCreditors = deSerializer.Deserialize(stream)
        stream.Close()

        '************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable
        Dim tempTable2 As DataTable

        If (tableSummaryOfPrice IsNot Nothing) Then
            If (tableSummaryOfPrice.Rows.Count > 0) Then
                view = New DataView(tableSummaryOfPrice)
                filter = "[Краткое описание] = 'Цена нетто в документе закупки в валюте документа' Or [Краткое описание] Like 'Новая%'"
                view.RowFilter = filter
                tableSummaryOfPrice = view.ToTable()

                'делаем по России
                view = New DataView(tableSummaryOfPrice)
                filter = "[БЕ] Not Like 'UA%'"
                view.RowFilter = filter
                tempTable = view.ToTable()
                'distinct Код кредитора
                view = New DataView(tempTable)
                tempTable = view.ToTable(True, "БЕ", "Кредитор")

                For i As Integer = 0 To tempTable.Rows.Count - 1
                    Dim row As DataRow = tempTable.Rows(i)
                    If Not DBNull.Value.Equals(row("Кредитор")) AndAlso row("Кредитор") <> Nothing Then
                        view = New DataView(tableBlocked)
                        filter = "[Счет] = '" & row("Кредитор") & "'"
                        view.RowFilter = filter
                        tempTable2 = view.ToTable()
                        If tempTable2.Rows.Count > 0 Then
                            For j As Integer = 0 To tempTable2.Rows.Count - 1
                                If row("БЕ") = tempTable2.Rows(j)("БЕ") Then
                                    Dim newRow As DataRow = tableBlocked2.NewRow()
                                    newRow("Счет") = tempTable2.Rows(j)("Счет")
                                    newRow("Наименование поставщика") = tempTable2.Rows(j)("Наименование поставщика")
                                    newRow("№ докум#") = tempTable2.Rows(j)("№ документа")
                                    newRow("ДатаДокум") = tempTable2.Rows(j)("ДатаДокум")
                                    newRow("Распределение мониторинга") = tempTable2.Rows(j)("Распределение мониторинга")
                                    tableBlocked2.Rows.Add(newRow)
                                End If
                            Next
                        End If
                    End If
                Next
                'проверка Исключения
                If (tableBlocked2 IsNot Nothing) Then
                    If (tableBlocked2.Rows.Count > 0) Then
                        For i As Integer = 0 To tableBlocked2.Rows.Count - 1
                            Dim row As DataRow = tableBlocked2.Rows(i)
                            If Not DBNull.Value.Equals(row("Счет")) AndAlso row("Счет") <> Nothing Then
                                view = New DataView(tableExceptionsForCreditors)
                                filter = "[Исключения для автопостинга] = '" & row("Счет") & "'"
                                view.RowFilter = filter
                                tempTable = view.ToTable()
                                If tempTable.Rows.Count > 0 Then
                                    row("исключение УП2") = tempTable.Rows(0)("Исключения для автопостинга")
                                Else
                                    row("исключение УП2") = "#Н/Д"
                                End If
                            End If
                        Next
                    End If
                End If


                'делаем по Украине
                view = New DataView(tableSummaryOfPrice)
                filter = "[БЕ] Like 'UA%'"
                view.RowFilter = filter
                tempTable = view.ToTable()
                'distinct Код кредитора
                view = New DataView(tempTable)
                tempTable = view.ToTable(True, "БЕ", "Кредитор")

                For i As Integer = 0 To tempTable.Rows.Count - 1
                    Dim row As DataRow = tempTable.Rows(i)
                    If Not DBNull.Value.Equals(row("Кредитор")) AndAlso row("Кредитор") <> Nothing Then
                        view = New DataView(tableBlockedUkr)
                        filter = "[Счет] = '" & row("Кредитор") & "'"
                        view.RowFilter = filter
                        tempTable2 = view.ToTable()
                        If tempTable2.Rows.Count > 0 Then
                            For j As Integer = 0 To tempTable2.Rows.Count - 1
                                If row("БЕ") = tempTable2.Rows(j)("БЕ") Then
                                    Dim newRow As DataRow = tableBlockedUkr2.NewRow()
                                    newRow("Счет") = tempTable2.Rows(j)("Счет")
                                    newRow("Наименование поставщика") = tempTable2.Rows(j)("Наименование поставщика")
                                    newRow("№ докум#") = tempTable2.Rows(j)("№ документа")
                                    newRow("ДатаДокум") = tempTable2.Rows(j)("ДатаДокум")
                                    newRow("Распределение мониторинга") = tempTable2.Rows(j)("Распределение мониторинга")
                                    tableBlockedUkr2.Rows.Add(newRow)
                                End If
                            Next
                        End If
                    End If
                Next

            End If
        End If



        Console.ReadKey()

    End Sub

End Module
