Imports Excel = Microsoft.Office.Interop.Excel

Public Class Excel_2007
    Public filePath As String = ""

    Public Sub Run()
        Try
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
            xlWorkSheet.Cells(1, 1) = "http://vb.net-informations.com"
            xlWorkSheet.SaveAs(filePath)
            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

        Catch ex As Exception

        End Try

        Me.Kill("EXCEL")

        Console.WriteLine("Excel file created , you can find the file c:\")

    End Sub

    Private Sub releaseObject(ByVal obj As Object)

        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Public Sub Kill(ByVal process_name As String)
        Try
            ' The excel is created and opened for insert value. We most close this excel using this system
            Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName(process_name.ToUpper())
            For Each i As Process In pro
                i.Kill()
            Next
        Catch ex As Exception

        End Try

    End Sub
End Class

