Imports Microsoft.Office.Interop

Module Module1

    Sub Main()
        Dim oExcel As Excel.Application
        Shell("C:\Program Files\Microsoft Office\Office\Excel.EXE", vbMinimizedNoFocus)    ' An error 429 occurs on the following line:    
        oExcel = GetObject(, "Excel.Application")
        oExcel.Workbooks.Item(0).Close()
        oExcel = Nothing
    End Sub


End Module
