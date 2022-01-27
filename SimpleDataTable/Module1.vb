Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()

        Dim table As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream("d:\Time\OB08_val.XML", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(table.GetType())

        table = deSerializer.Deserialize(stream)
        stream.Close()

        '************** Begin
        Dim view As DataView
        Dim filter As String
        Dim column As String
        Dim tempTable As DataTable
        Dim isRowPresent As Boolean

        tempTable = table.Clone()
        For i As Integer = 0 To 2
            Dim row As DataRow = table.Rows(i)
            tempTable.ImportRow(row)
        Next

        stream = New FileStream("d:\Time\1_OB08_val.XML", FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(tempTable.GetType())
        serializer.Serialize(stream, tempTable)
        stream.Close()

        column = "исключение УП2"
        filter = "#Н/Д"
        isRowPresent = False

        view = New DataView(table)
        filter = "[" & column & "] <> '" & filter & "'"
        view.RowFilter = filter
        tempTable = view.ToTable()

        If (tempTable.Rows.Count > 0) Then
            isRowPresent = True
        End If

        Console.ReadKey()

    End Sub

End Module
