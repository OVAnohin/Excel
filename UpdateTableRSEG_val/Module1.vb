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
        Dim tableOB08 As DataTable = New DataTable()
        stream = New FileStream("d:\Time\OB08.xml", FileMode.Open, FileAccess.Read)
        deSerializer = New XmlSerializer(tableOB08.GetType())

        tableOB08 = deSerializer.Deserialize(stream)
        stream.Close()

        '**************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable

        If (tableRSEG IsNot Nothing) Then
            If (tableRSEG.Rows.Count > 0) Then
                tableRSEG.Columns.Add("Валюта", Type.GetType("System.String"))

                For i As Integer = 0 To tableRSEG.Rows.Count - 1
                    Dim row As DataRow = tableRSEG.Rows(i)
                    If Not DBNull.Value.Equals(row("ДокумЗакуп")) AndAlso row("ДокумЗакуп") <> Nothing AndAlso row("ДокумЗакуп") <> "" Then
                        view = New DataView(tableOB08)
                        Dim searchString As String = row("ДокумЗакуп")
                        filter = "[ДокумЗакуп] = '" & searchString & "'"
                        view.RowFilter = filter
                        tempTable = view.ToTable()
                        If tempTable.Rows.Count > 0 Then
                            Dim val As String = tempTable.Rows(0)("Влт")
                            Select Case val
                                Case "RUE"
                                    row("Валюта") = "EUR"
                                Case "RUD"
                                    row("Валюта") = "USD"
                                Case Else
                                    row("Валюта") = "RUB"
                            End Select
                        Else
                            row("Валюта") = "NotFound"
                        End If
                    End If
                Next
            End If
        End If
    End Sub

End Module
