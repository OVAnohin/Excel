Module Module1

    Sub Main()

        Dim tempTable As New DataTable("")

        tempTable.Columns.Add("DateC", Type.GetType("System.DateTime"))
        tempTable.Columns.Add("DatePo", Type.GetType("System.DateTime"))
        tempTable.Columns.Add("YesOrNo", Type.GetType("System.String"))

        If (tempTable IsNot Nothing) Then
            If (tempTable.Rows.Count > 0) Then
                Console.WriteLine(tempTable.Rows.Count)
                Console.ReadKey()
            End If
        End If

        Console.ReadKey()

    End Sub

End Module
