Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()

        Dim tableNewContracts As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\NewСontracts.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tableNewContracts.GetType())

        tableNewContracts = deSerializer.Deserialize(stream)
        stream.Close()

        Dim tableSummaryOfContracts As DataTable
        Dim view As DataView

        If (tableNewContracts IsNot Nothing) Then
            If (tableNewContracts.Rows.Count > 0) Then
                view = New DataView(tableNewContracts)
                view.Sort = "БЕ, ДокумЗакуп, Поставщик, Наименование, ДатаСоздан, НачальнС, Конец времени выполнения"
                tableSummaryOfContracts = view.ToTable(False, "БЕ", "ДокумЗакуп", "Поставщик", "Наименование", "ДатаСоздан", "НачальнС", "Конец времени выполнения")
            End If
        End If


        Console.ReadKey()
    End Sub

End Module
