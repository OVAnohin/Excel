Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        Dim tableRSEG As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\RSEG.XML", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableRSEG.GetType())

        tableRSEG = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableBlocked.xml
        Dim tableBlocked2 As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableBlocked2.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableBlocked2.GetType())

        tableBlocked2 = deSerializer.Deserialize(stream)
        stream.Close()

        '**************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        If (tableRSEG IsNot Nothing) Then
            If (tableRSEG.Rows.Count > 0) Then
                tableRSEG.Columns.Add("Счет", Type.GetType("System.String")).SetOrdinal(1)
                tableRSEG.Columns.Add("Цена", Type.GetType("System.Decimal"))
                tableRSEG.Columns.Add("Дата документа", Type.GetType("System.String"))
                tableRSEG.Columns.Add("ME3L (UA_SAP)", Type.GetType("System.String"))
                tableRSEG.Columns.Add("проверка Вариант1", Type.GetType("System.String"))
                tableRSEG.Columns.Add("ME3L (VLALI5)", Type.GetType("System.String"))
                tableRSEG.Columns.Add("проверка Вариант2", Type.GetType("System.String"))
                tableRSEG.Columns.Add("курс из OB08", Type.GetType("System.String"))
                tableRSEG.Columns.Add("Контракт", Type.GetType("System.String"))
                tableRSEG.Columns.Add("Закупщик", Type.GetType("System.String"))
                tableRSEG.Columns.Add("Условие платежа", Type.GetType("System.String"))

                For i As Integer = 0 To tableRSEG.Rows.Count - 1
                    Dim row As DataRow = tableRSEG.Rows(i)
                    If Not DBNull.Value.Equals(row("BELNR")) AndAlso row("BELNR") <> Nothing AndAlso row("BELNR") <> "" Then
                        view = New DataView(tableBlocked2)
                        Dim searchString As String = row("BELNR")
                        filter = "[№ докум#] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            row("Счет") = tempTable.Rows(0)("Счет")
                            row("Дата документа") = tempTable.Rows(0)("ДатаДокум")
                        Else
                            row("Счет") = ""
                            row("Дата документа") = ""
                        End If
                        'row("Цена") = CInt(Int((row("WRBTR") / row("BPMNG")) * 100)) / 100
                        row("Цена") = Math.Round((row("WRBTR") / row("BPMNG")), 2)
                    End If
                Next
            End If

            tableRSEG.Columns("BELNR").ColumnName = "№ докум"
            tableRSEG.Columns("EBELN").ColumnName = "ДокумЗакуп"
            tableRSEG.Columns("EBELP").ColumnName = "Поз"
            tableRSEG.Columns("MATNR").ColumnName = "Материал"
            tableRSEG.Columns("BUKRS").ColumnName = "БЕ"
            tableRSEG.Columns("WERKS").ColumnName = "З-д"
            tableRSEG.Columns("WRBTR").ColumnName = "Сумма"
            tableRSEG.Columns("BPMNG").ColumnName = "Колич/ЕИЦЗ"
        End If


    End Sub

End Module
