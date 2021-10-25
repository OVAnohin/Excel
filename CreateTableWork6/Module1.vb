Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        'report
        Dim tableReport As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\Report.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableReport.GetType())

        tableReport = deSerializer.Deserialize(stream)
        stream.Close()

        Dim tableWork6 As DataTable

        '************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        'Filter = Новая позиция*, Цена нетто в документе закупки в валюте документа
        view = New DataView(tableReport)
        filter = "[Краткое описание] = 'Цена нетто в документе закупки в валюте документа' OR [Краткое описание] Like 'Новая позиция%'"
        view.RowFilter = filter
        tableWork6 = view.ToTable()
        tableWork6.Columns.Add("Сцепить", Type.GetType("System.String"))

        For i As Integer = 0 To tableWork6.Rows.Count - 1
            Dim row As DataRow = tableWork6.Rows(i)
            If Not DBNull.Value.Equals(row("Номер Договора")) AndAlso Not DBNull.Value.Equals(row("Позиция Контракта")) Then
                row("Сцепить") = row("Номер Договора") & row("Позиция Контракта")
                Continue For
            End If
            If Not DBNull.Value.Equals(row("Номер Договора")) AndAlso DBNull.Value.Equals(row("Позиция Контракта")) Then
                row("Сцепить") = row("Номер Договора")
                Continue For
            End If
            If DBNull.Value.Equals(row("Номер Договора")) AndAlso Not DBNull.Value.Equals(row("Позиция Контракта")) Then
                row("Сцепить") = row("Позиция Контракта")
                Continue For
            End If
            row("Сцепить") = ""
        Next

        view = New DataView(tableWork6)
        tableWork6 = view.ToTable(False, "Номер Договора", "Сцепить")

        'часть 2
        'report
        Dim tableMaterials As DataTable = New DataTable()
        stream = New FileStream("d:\Time\Materials.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableReport.GetType())

        tableMaterials = deSerializer.Deserialize(stream)
        stream.Close()

        tableMaterials.Columns.Add("Сцепить", Type.GetType("System.String"))
        For i As Integer = 0 To tableMaterials.Rows.Count - 1
            Dim row As DataRow = tableMaterials.Rows(i)
            If Not DBNull.Value.Equals(row("EBELN")) AndAlso Not DBNull.Value.Equals(row("EBELP")) Then
                row("Сцепить") = row("EBELN") & row("EBELP")
                Continue For
            End If
            If Not DBNull.Value.Equals(row("EBELN")) AndAlso DBNull.Value.Equals(row("EBELP")) Then
                row("Сцепить") = row("EBELN")
                Continue For
            End If
            If DBNull.Value.Equals(row("EBELN")) AndAlso Not DBNull.Value.Equals(row("EBELP")) Then
                row("Сцепить") = row("EBELP")
                Continue For
            End If
            row("Сцепить") = ""
        Next

        For i As Integer = 0 To tableReport.Rows.Count - 1
            Dim row As DataRow = tableReport.Rows(i)
            'filter = "[Краткое описание] = 'Цена нетто в документе закупки в валюте документа' OR [Краткое описание] Like 'Новая позиция%'"
            If Not DBNull.Value.Equals(row("Краткое описание")) AndAlso row("Краткое описание") <> Nothing Then
                If Left(row("Краткое описание"), 13) = "Новая позиция" Or row("Краткое описание") = "Цена нетто в документе закупки в валюте документа" Then
                    Dim strMaterial As String = ""
                    If Not DBNull.Value.Equals(row("Номер Договора")) AndAlso Not DBNull.Value.Equals(row("Позиция Контракта")) Then
                        strMaterial = row("Номер Договора") & row("Позиция Контракта")
                    ElseIf Not DBNull.Value.Equals(row("Номер Договора")) AndAlso DBNull.Value.Equals(row("Позиция Контракта")) Then
                        strMaterial = row("Номер Договора")
                    ElseIf DBNull.Value.Equals(row("Номер Договора")) AndAlso Not DBNull.Value.Equals(row("Позиция Контракта")) Then
                        strMaterial = row("Позиция Контракта")
                    End If

                    view = New DataView(tableMaterials)
                    filter = "[Сцепить] = '" & strMaterial & "'"
                    view.RowFilter = filter
                    tempTable = view.ToTable()
                    If tempTable.Rows.Count > 0 Then
                        row("Материал") = tempTable.Rows(0)("MATNR").ToString()
                    Else
                        row("Материал") = ""
                    End If
                End If
            End If
        Next

        Console.ReadKey()
    End Sub

End Module
