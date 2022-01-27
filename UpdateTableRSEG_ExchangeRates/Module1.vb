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
        Dim tableOB08_val As DataTable = New DataTable()
        stream = New FileStream("d:\Time\OB08_val.XML", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableOB08_val.GetType())

        tableOB08_val = deSerializer.Deserialize(stream)
        stream.Close()

        '**************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        If (tableRSEG IsNot Nothing) Then
            If (tableRSEG.Rows.Count > 0) Then
                For i As Integer = 0 To tableRSEG.Rows.Count - 1
                    Dim row As DataRow = tableRSEG.Rows(i)
                    If Not DBNull.Value.Equals(row("Валюта")) AndAlso row("Валюта") <> Nothing AndAlso row("Валюта") <> "" AndAlso row("Валюта") <> "NotFound" Then
                        view = New DataView(tableOB08_val)
                        Dim searchString As String = row("Дата документа")
                        filter = "[Действит# с] = '" & Left(searchString, 10) & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            For tRow As Integer = 0 To tempTable.Rows.Count - 1
                                Dim exchange As Double = tempTable.Rows(tRow)("Курс")
                                If row("Валюта") = tempTable.Rows(tRow)("Исходная валюта") Then
                                    row("Цена") = Math.Round((row("Сумма") / row("Колич/ЕИЦЗ") / tempTable.Rows(tRow)("Курс")), 2)
                                    row("курс из OB08") = tempTable.Rows(tRow)("Курс").ToString()
                                    Exit For
                                End If
                            Next
                        Else
                            row("Цена") = 0
                        End If
                    End If
                Next
            End If
        End If
    End Sub

End Module
