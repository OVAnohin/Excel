Imports System.Data.OleDb

Module Module1

    Sub Main()

        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\TablePaymentTermsCodeFalset.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Sheet1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim TablePaymentTermsCodeFalse As DataTable
        oDataAdapter.Fill(oDataSet)
        TablePaymentTermsCodeFalse = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        'table work1
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\ZRUV_ECM_PT_Filtered.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        sheetName = "Sheet1"
        oDataAdapter = New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        oDataSet = New DataSet
        Dim ZRUV_ECM_PT_Filtered As DataTable
        oDataAdapter.Fill(oDataSet)
        ZRUV_ECM_PT_Filtered = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        '**************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        TablePaymentTermsCodeFalse.Columns.Add("DateC", Type.GetType("System.DateTime"))
        TablePaymentTermsCodeFalse.Columns.Add("DatePo", Type.GetType("System.DateTime"))
        TablePaymentTermsCodeFalse.Columns.Add("YesOrNo", Type.GetType("System.String"))

        For i As Integer = 0 To TablePaymentTermsCodeFalse.Rows.Count - 1
            Dim row As DataRow = TablePaymentTermsCodeFalse.Rows(i)
            Dim searchString As String = row("Ссылка на платеж")
            Dim rowAI As String = row("AI")
            view = New DataView(ZRUV_ECM_PT_Filtered)
            filter = "[Номер контракта] = " & row("Ссылка на платеж")
            view.RowFilter = filter
            tempTable = view.ToTable()
            If tempTable.Rows.Count = 0 Then
                row("DateC") = ""
                row("DatePo") = ""
                row("YesOrNo") = "да"
                Continue For
            End If

            Dim rowFiltered As DataRow = tempTable.Rows(0)

            If Not DBNull.Value.Equals(rowFiltered("Условие платежа 3")) AndAlso rowFiltered("Условие платежа 3") <> Nothing AndAlso rowFiltered("Условие платежа 3") = row("AI") Then
                row("DateC") = Left(rowFiltered("Дата начала действия условий платежа 3").ToString(), 10)
                row("DatePo") = Left(rowFiltered("Дата окончания действия условий платежа 3").ToString(), 10)
                If row("Дата документа") >= rowFiltered("Дата начала действия условий платежа 3") AndAlso row("Дата документа") <= rowFiltered("Дата окончания действия условий платежа 3") Then
                    row("YesOrNo") = "да"
                Else
                    row("YesOrNo") = "нет"
                End If
                Continue For
            End If
            If Not DBNull.Value.Equals(rowFiltered("Условие платежа 2")) AndAlso rowFiltered("Условие платежа 2") <> Nothing AndAlso rowFiltered("Условие платежа 2") = row("AI") Then
                row("DateC") = rowFiltered("Дата начала действия условий платежа 2")
                row("DatePo") = rowFiltered("Дата окончания действия условий платежа 2")
                If row("Дата документа") >= rowFiltered("Дата начала действия условий платежа 2") AndAlso row("Дата документа") <= rowFiltered("Дата окончания действия условий платежа 2") Then
                    row("YesOrNo") = "да"
                Else
                    row("YesOrNo") = "нет"
                End If
                Continue For
            End If
            If Not DBNull.Value.Equals(rowFiltered("Условие платежа 1")) AndAlso rowFiltered("Условие платежа 1") <> Nothing AndAlso rowFiltered("Условие платежа 1") = row("AI") Then
                row("DateC") = Left(rowFiltered("Дата начала действия условий платежа 1").ToString(), 10)
                row("DatePo") = rowFiltered("Дата окончания действия условий платежа")
                If row("Дата документа") >= rowFiltered("Дата начала действия условий платежа 1") AndAlso row("Дата документа") <= rowFiltered("Дата окончания действия условий платежа") Then
                    row("YesOrNo") = "да"
                Else
                    row("YesOrNo") = "нет"
                End If
                Continue For
            End If
            row("YesOrNo") = "да"
        Next

        TablePaymentTermsCodeFalse.Columns.Add("Сцепить", Type.GetType("System.String"))
        TablePaymentTermsCodeFalse.Columns.Add("Контракт", Type.GetType("System.String"))

        'СЦЕПИТЬ(J2;"_";N2)
        '#Н/Д

        For i As Integer = 0 To TablePaymentTermsCodeFalse.Rows.Count - 1
            Dim row As DataRow = TablePaymentTermsCodeFalse.Rows(i)
            row("Сцепить") = row("Блк") & "_" & row("Текст заголовка документа")
        Next

        For i As Integer = 0 To TablePaymentTermsCodeFalse.Rows.Count - 1
            Dim row As DataRow = TablePaymentTermsCodeFalse.Rows(i)
            view = New DataView(ZRUV_ECM_PT_Filtered)

            If row("СсылкаПлат") <> Nothing AndAlso row("СсылкаПлат") <> "" Then
                filter = "[Номер Договора] = " & row("СсылкаПлат")
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
        tableParkedBlocked = TablePaymentTermsCodeFalse.Clone()
        Dim tableTemp2 As DataTable
        tableTemp2 = TablePaymentTermsCodeFalse.Clone()
        Dim tableTemp3 As DataTable
        tableTemp3 = TablePaymentTermsCodeFalse.Clone()

        'Filter = "@5C\Qоткрыт.@"
        view = New DataView(TablePaymentTermsCodeFalse)
        filter = "[Ст] = '@5C\Qоткрыт.@'"
        view.RowFilter = filter
        Dim tableOpenPosition As DataTable = view.ToTable()

        'Filter = "@5D\QПредвРег@"
        view = New DataView(TablePaymentTermsCodeFalse)
        filter = "[Ст] = '@5D\QПредвРег@'"
        view.RowFilter = filter
        Dim tablePreRegistration As DataTable = view.ToTable()

        '(19-58) Фильтр по столбцу А по тексту «@5C\Qоткрыт.@»
        '                                       @5C\Qоткрыт.@
        'Фильтр (20-52) по столбцу AJ «Сцепить»: сначала по значению «Х*», затем по значению «W*,01*»
        view = New DataView(tableOpenPosition)
        filter = "[Сцепить] Like 'Х%'"
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
        filter = "[Контракт] Like '#Н/Д' AND [СсылкаПлат] = ''"
        view.RowFilter = filter
        tempTable = view.ToTable()
        For i As Integer = 0 To tempTable.Rows.Count - 1
            tableParkedBlocked.ImportRow(tempTable.Rows(i))
        Next


        '************************************************
        'Фильтр по столбцу А по тексту «@5D\QПредвРег@». Фильтр по столбцу N «Текст заголовка документа» по значению «,01».
        view = New DataView(tablePreRegistration)
        filter = "[Текст заголовка документа] Like '%,01%'"
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
            view = New DataView(ZRUV_ECM_PT_Filtered)

            If row("Счет") <> Nothing AndAlso row("Счет") <> "" Then
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

        ShowTable(tableParkedBlocked)
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
