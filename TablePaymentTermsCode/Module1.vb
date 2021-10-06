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

        'table PaymentTermsCode
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\PaymentTermsCode.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        sheetName = "Sheet1"
        oDataAdapter = New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        oDataSet = New DataSet
        Dim PaymentTermsCode As DataTable
        oDataAdapter.Fill(oDataSet)
        PaymentTermsCode = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        'table Work2
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Time\Work2.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        sheetName = "Sheet1"
        oDataAdapter = New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        oDataSet = New DataSet
        Dim Work2 As DataTable
        oDataAdapter.Fill(oDataSet)
        Work2 = oDataSet.Tables(0)
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

        PaymentTermsCode.Columns.Add("DateC", Type.GetType("System.String"))
        PaymentTermsCode.Columns.Add("DatePo", Type.GetType("System.String"))
        PaymentTermsCode.Columns.Add("YesOrNo", Type.GetType("System.String"))

        Dim nullTable As DataTable
        view = New DataView(PaymentTermsCode)
        filter = "CONVERT(Isnull([Ссылка на платеж],''), System.String) = ''"
        view.RowFilter = filter
        nullTable = view.ToTable()

        For i As Integer = 0 To PaymentTermsCode.Rows.Count - 1
            Dim row As DataRow = PaymentTermsCode.Rows(i)
            view = New DataView(nullTable)
            filter = "[№ документа] = " & row("№ документа") & " and [Дата документа] = '" & row("Дата документа") & "'" & " and [Дата проводки] = '" & row("Дата проводки") & "'"
            view.RowFilter = filter
            tempTable = view.ToTable()
            If tempTable.Rows.Count > 0 Then
                row("Ссылка на платеж") = "0" ' На призме тут "''"
                row("YesOrNo") = "да"
                Continue For
            End If
        Next

        For i As Integer = 0 To PaymentTermsCode.Rows.Count - 1
            Dim row As DataRow = PaymentTermsCode.Rows(i)
            view = New DataView(TablePaymentTermsCodeFalse)
            'filter = "[Ссылка на платеж] = " & row("Ссылка на платеж") & " and [№ документа] = " & row("№ документа") & "and [Дата документа] = " & row("Дата документа") & "And [Дата проводки] = " & row("Дата проводки")
            filter = "[№ документа] = " & row("№ документа") & " and [Дата документа] = '" & row("Дата документа") & "'" & " and [Дата проводки] = '" & row("Дата проводки") & "'"
            view.RowFilter = filter
            tempTable = view.ToTable()
            If tempTable.Rows.Count > 0 Then
                row("DateC") = tempTable.Rows(0)("DateC")
                row("DatePo") = tempTable.Rows(0)("DatePo")
                row("YesOrNo") = tempTable.Rows(0)("YesOrNo")
                Continue For
            End If
        Next

        Dim VerificationTab As DataTable
        VerificationTab = PaymentTermsCode.Clone()

        view = New DataView(PaymentTermsCode)
        filter = "CONVERT(Isnull([Ссылка на платеж],''), System.String) <> ''"
        view.RowFilter = filter
        VerificationTab = view.ToTable()


        view = New DataView(PaymentTermsCode)
        filter = "[YesOrNo] = 'да'"
        view.RowFilter = filter
        VerificationTab = view.ToTable()

        VerificationTab.Columns.Remove("DateC")
        VerificationTab.Columns.Remove("DatePo")
        VerificationTab.Columns.Remove("YesOrNo")

        VerificationTab.Columns.Add("ЧЕК", Type.GetType("System.String"))
        VerificationTab.Columns.Add("Изменения внесены да/нет", Type.GetType("System.String"))
        VerificationTab.Columns.Add("комментарии", Type.GetType("System.String"))
        VerificationTab.Columns.Add("новое", Type.GetType("System.String"))
        VerificationTab.Columns.Add("старое", Type.GetType("System.String"))
        VerificationTab.Columns.Add("контракт", Type.GetType("System.String"))

        For i As Integer = 0 To VerificationTab.Rows.Count - 1
            Dim row As DataRow = VerificationTab.Rows(i)
            If DBNull.Value.Equals(row("Ссылка на платеж")) Then
                row("контракт") = "0"
                Continue For
            End If
            row("контракт") = row("Ссылка на платеж")
            view = New DataView(Work2)
            filter = "[Номер Договора] = " & row("Ссылка на платеж")
            view.RowFilter = filter
            tempTable = view.ToTable()
            If tempTable.Rows.Count > 0 Then
                row("новое") = tempTable.Rows(0)("Новое значение")
                row("старое") = tempTable.Rows(0)("Старое значение")
                row("ЧЕК") = tempTable.Rows(0)("Краткое описание")
                Continue For
            End If
        Next


        ShowTable(VerificationTab)

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
