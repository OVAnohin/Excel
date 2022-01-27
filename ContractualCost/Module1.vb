Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()

        Dim tableContractualCost As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\TableContractualCost.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableContractualCost.GetType())

        tableContractualCost = deSerializer.Deserialize(stream)
        stream.Close()

        'table TableBlocked.xml
        Dim tableWork1 As DataTable = New DataTable()
        stream = New FileStream("d:\Time\TableWork1.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableWork1.GetType())

        tableWork1 = deSerializer.Deserialize(stream)
        stream.Close()

        '**************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        For i As Integer = 0 To tableContractualCost.Rows.Count - 1
            Dim row As DataRow = tableContractualCost.Rows(i)
            If Not DBNull.Value.Equals(row("СсылкаПлат")) Then
                Dim searchString As String = row("СсылкаПлат")
                Dim searchChar = "/"
                Dim position = InStr(1, searchString, searchChar, 1)
                If position <> Nothing AndAlso position <> 0 Then
                    row("СсылкаПлат") = Mid(searchString, position + 1)
                End If
            End If
        Next

        For i As Integer = 0 To tableContractualCost.Rows.Count - 1
            Dim row As DataRow = tableContractualCost.Rows(i)
            Dim searchString As String
            Dim contractNumber As String
            If Not DBNull.Value.Equals(row("СсылкаПлат")) Then
                searchString = row("СсылкаПлат")
            Else
                searchString = Nothing
            End If
            If Not DBNull.Value.Equals(row("№ договора")) Then
                contractNumber = row("№ договора")
            Else
                contractNumber = Nothing
            End If
            If DBNull.Value.Equals(row("СсылкаПлат")) AndAlso DBNull.Value.Equals(row("№ договора")) Then
                Continue For
            End If
            If searchString = Nothing AndAlso contractNumber = Nothing Then
                Continue For
            End If
            If searchString = Nothing AndAlso contractNumber <> Nothing Then
                row("СсылкаПлат") = row("№ договора")
                Continue For
            End If
            If searchString <> Nothing AndAlso Left(searchString, 2) <> "46" Then
                row("СсылкаПлат") = row("№ договора")
            End If
        Next

        tableContractualCost.Columns.Add("Сцепить", Type.GetType("System.String"))
        tableContractualCost.Columns.Add("Контракт", Type.GetType("System.String"))

        'СЦЕПИТЬ(J2;"_";N2)
        '#Н/Д

        For i As Integer = 0 To tableContractualCost.Rows.Count - 1
            Dim row As DataRow = tableContractualCost.Rows(i)
            row("Сцепить") = row("Блк") & "_" & row("Текст заголовка документа")
        Next

        For i As Integer = 0 To tableContractualCost.Rows.Count - 1
            Dim row As DataRow = tableContractualCost.Rows(i)
            view = New DataView(tableWork1)

            If Not DBNull.Value.Equals(row("СсылкаПлат")) AndAlso row("СсылкаПлат") <> Nothing AndAlso row("СсылкаПлат").ToString() <> "" Then
                filter = "[Номер Договора] = '" & row("СсылкаПлат") & "'"
                view.RowFilter = filter
                tempTable = view.ToTable()
                If tempTable.Rows.Count > 0 Then
                    row("Контракт") = tempTable.Rows(0)("Номер Договора")
                Else
                    row("Контракт") = "#Н/Д"
                End If
            Else
                row("Контракт") = "#Н/Д"
            End If
        Next

        Dim tableParkedBlocked As DataTable
        tableParkedBlocked = tableContractualCost.Clone()
        Dim tableTemp2 As DataTable
        tableTemp2 = tableContractualCost.Clone()
        Dim tableTemp3 As DataTable
        tableTemp3 = tableContractualCost.Clone()

        'Filter = "@5C\Qоткрыт.@"
        view = New DataView(tableContractualCost)
        filter = "[Ст] = '@5C\Qоткрыт.@'"
        view.RowFilter = filter
        Dim tableOpenPosition As DataTable = view.ToTable()

        'Filter = "@5D\QПредвРег@"
        view = New DataView(tableContractualCost)
        filter = "[Ст] = '@5D\QПредвРег@'"
        view.RowFilter = filter
        Dim tablePreRegistration As DataTable = view.ToTable()

        '(19-58) Фильтр по столбцу А по тексту «@5C\Qоткрыт.@»
        'Фильтр (20-52) по столбцу AJ «Сцепить»: сначала по значению «Х*», затем по значению «W*,01*»
        view = New DataView(tableOpenPosition)
        filter = "[Сцепить] Like 'X%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp2.ImportRow(tempTable.Rows(i))
        Next

        'затем по значению «W*,01*» добавляем таблицу tableTemp3 !!!!
        view = New DataView(tableOpenPosition)
        filter = "[Сцепить] Like 'W%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp3.ImportRow(tempTable.Rows(i))
        Next

        view = New DataView(tableTemp3)
        filter = "[Сцепить] Like '%,01%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableTemp2.ImportRow(tempTable.Rows(i))
        Next

        'Фильтр по столбцу AK «Контракт» по значению «46*». Отфильтрованные данные скопировать на вкладку «запарк|заблок».
        view = New DataView(tableTemp2)
        filter = "[Контракт] Like '46%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        'Фильтр по столбцу АК «Контракт» по значению «#Н/Д», фильтр по столбцу В по пустым. Отфильтрованные данные скопировать на вкладку «запарк|заблок».
        view = New DataView(tableTemp2)
        'filter = "[Контракт] Like '#Н/Д' AND [СсылкаПлат] = ''"
        filter = "[Контракт] Like '#Н/Д' AND [СсылкаПлат] IS NULL OR [СсылкаПлат] = ''"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        '************************************************
        '!!!! Old Фильтр по столбцу А по тексту «@5D\QПредвРег@». Фильтр по столбцу N «Текст заголовка документа» по значению «,01».

        '1)	Фильтр по столбцу Y «БЕ» по значению «RU*»
        'Фильтр по столбцу N «Текст заголовка документа» по значению «,01».
        view = New DataView(tablePreRegistration)
        filter = "[БЕ] Like 'RU%' AND [Текст заголовка документа] Like '%,01'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        '2)	Фильтр по столбцу Y «БЕ» по значению «UA01»
        'Фильтр по столбцу N «Текст заголовка документа» по значению «,01»
        view = New DataView(tablePreRegistration)
        filter = "[БЕ] = 'UA01' AND [Текст заголовка документа] Like '%,01'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        'Фильтр по столбцу N «Текст заголовка документа» по значению «,05»
        view = New DataView(tablePreRegistration)
        filter = "[БЕ] = 'UA01' AND [Текст заголовка документа] Like '%,05'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        'Фильтр по столбцу N «Текст заголовка документа» по значению «,05,»
        view = New DataView(tablePreRegistration)
        filter = "[БЕ] = 'UA01' AND [Текст заголовка документа] Like '%,05,%'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        'Фильтр по столбцу N «Текст заголовка документа» по значению «,07B» (В – латинская)
        view = New DataView(tablePreRegistration)
        filter = "[БЕ] = 'UA01' AND [Текст заголовка документа] Like '%,07B'"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next

        tableParkedBlocked.Columns.Add("Новое значение", Type.GetType("System.String"))
        tableParkedBlocked.Columns.Add("Старое значение", Type.GetType("System.String"))
        tableParkedBlocked.Columns.Add("Валюта", Type.GetType("System.String"))
        tableParkedBlocked.Columns.Add("Договор", Type.GetType("System.String"))
        tableParkedBlocked.Columns.Add("Чек", Type.GetType("System.String"))

        For i As Integer = 0 To tableParkedBlocked.Rows.Count - 1
            Dim row As DataRow = tableParkedBlocked.Rows(i)
            view = New DataView(tableWork1)

            If Not DBNull.Value.Equals(row("Счет")) AndAlso row("Счет") <> Nothing AndAlso row("Счет").ToString() <> "" Then
                filter = "[Кредитор] = " & row("Счет")
                view.RowFilter = filter
                tempTable = view.ToTable()
                If tempTable.Rows.Count > 0 Then
                    row("Новое значение") = tempTable.Rows(0)("Новое значение")
                    row("Старое значение") = tempTable.Rows(0)("Старое значение")
                    row("Валюта") = tempTable.Rows(0)("Валюта")
                    row("Договор") = tempTable.Rows(0)("Номер Договора")
                    row("Чек") = tempTable.Rows(0)("Краткое описание")
                Else
                    row("Новое значение") = "#Н/Д"
                    row("Старое значение") = "#Н/Д"
                    row("Валюта") = "#Н/Д"
                    row("Договор") = "#Н/Д"
                    row("Чек") = "#Н/Д"
                End If
            Else
                row("Новое значение") = "#Н/Д"
                row("Старое значение") = "#Н/Д"
                row("Валюта") = "#Н/Д"
                row("Договор") = "#Н/Д"
                row("Чек") = "#Н/Д"
            End If
        Next

        tableParkedBlocked.Columns.Remove("Сцепить")
        tableParkedBlocked.Columns.Remove("Контракт")

        'tableParkedBlockedOut = tableParkedBlocked
        'tableContractualCostOut = tableContractualCost

        ShowTable(tableContractualCost)
        Console.ReadKey()

    End Sub

    Private Sub ShowTable(table As DataTable)
        For i = 0 To table.Rows.Count - 1
            For j = 0 To table.Columns.Count - 1
                Dim row As DataRow = table.Rows(i)
                Dim column As DataColumn = table.Columns(j)
                Console.Write(row(column) & " ")
            Next
            Console.WriteLine()
        Next
        Console.WriteLine(New String("*", 20))
    End Sub

End Module
