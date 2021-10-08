Imports System.Data.OleDb
Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()
        Dim localFolder As String = "d:\Work"
        Dim xmlFileName As String = "Datatable.xml"
        Dim oConnection As OleDbConnection
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\Work\123.xlsx;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If oConnection Is Nothing Then
            oConnection = New OleDbConnection(connectionString)
            oConnection.Open()
        End If

        Dim sheetName As String = "Sheet1"
        'sheetName = "Sheet1"
        Dim oDataAdapter As New OleDbDataAdapter("Select * from[" & sheetName & "$]", oConnection)
        Dim oDataSet As New DataSet
        Dim dataTable As DataTable
        oDataAdapter.Fill(oDataSet)
        dataTable = oDataSet.Tables(0)
        oConnection.Close()
        oConnection = Nothing

        ShowTable(dataTable)

        'Using stream As FileStream = New FileStream(localFolder & "\" & xmlFileName, FileMode.Create)
        '    Dim serializer As XmlSerializer = New XmlSerializer(dataTable.GetType())
        '    serializer.Serialize(stream, dataTable)
        'End Using

        Dim stream As FileStream = New FileStream(localFolder & "\" & xmlFileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(dataTable.GetType())

        serializer.Serialize(stream, dataTable)
        stream.Close()

        Dim newTable As DataTable
        stream = New FileStream(localFolder & "\" & xmlFileName, FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(dataTable.GetType())
        newTable = deSerializer.Deserialize(stream)
        stream.Close()

        'Try
        '    Using fsSource As FileStream = New FileStream(localFolder & "\" & xmlFileName, FileMode.Open, FileAccess.Read)
        '        ' Read the source file into a byte array.
        '        Dim deSerializer As XmlSerializer = New XmlSerializer(dataTable.GetType())


        '    End Using
        'Catch ioEx As FileNotFoundException
        '    Console.WriteLine(ioEx.Message)
        'End Try

        'dataTable.WriteXml(localFolder & "\" & xmlFileName, True)

        ''Dim writer As New IO.StringWriter
        ''dataTable.WriteXml(writer, XmlWriteMode.WriteSchema, False)
        ''PrintOutput(writer, "dataTable, without hierarchy")

        ''writer = New IO.StringWriter
        ''dataTable.WriteXml(writer, XmlWriteMode.WriteSchema, True)
        ''PrintOutput(writer, "dataTable, with hierarchy")

        'Dim newTable As DataTable
        'newTable.ReadXml(localFolder & "\" & xmlFileName)

        Console.WriteLine("Press any key to continue.")
        Console.ReadKey()

    End Sub

    Private Sub PrintOutput(ByVal writer As IO.TextWriter, ByVal caption As String)
        Console.WriteLine("==============================")
        Console.WriteLine(caption)
        Console.WriteLine("==============================")
        Console.WriteLine(writer.ToString())
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
