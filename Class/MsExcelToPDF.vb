Imports Excel = Microsoft.Office.Interop.Excel

'http:'stackoverflow.com/questions/5499562/excel-to-pdf-c-sharp-library
Public Class MsExcelToPDF

    Sub New()

    End Sub

    Public Function ExportWorkbookToPdf(ByVal workbookPath As String, ByVal outputPath As String) As Boolean
        ' If either required string is null or empty, stop and bail out
        If (String.IsNullOrEmpty(workbookPath) Or String.IsNullOrEmpty(outputPath)) Then
            Return False
        End If
        ' Create COM Objects
        Dim excelApplication As Microsoft.Office.Interop.Excel.Application
        Dim excelWorkbook As Microsoft.Office.Interop.Excel.Workbook

        ' Create new instance of Excel
        excelApplication = New Microsoft.Office.Interop.Excel.Application()

        ' Make the process invisible to the user
        excelApplication.ScreenUpdating = False

        ' Make the process silent
        excelApplication.DisplayAlerts = False

        ' Open the workbook that you wish to export to PDF
        excelWorkbook = excelApplication.Workbooks.Open(workbookPath)

        ' If the workbook failed to open, stop, clean up, and bail out
        If False Then
            excelApplication.Quit()

            excelApplication = Nothing
            excelWorkbook = Nothing

            Return False
        End If

        Dim exportSuccessful As Boolean = True
        Try
            'excelWorkbook.ActiveSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape

            ' Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
            excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath)
        Catch ex As Exception
            ' Mark the export as failed for the return value...
            exportSuccessful = False

            ' Do something with any exceptions here, if you wish...
            ' MessageBox.Show...        
        Finally
            ' Close the workbook, quit the Excel, and clean up regardless of the results...
            excelWorkbook.Close()
            excelApplication.Quit()

            excelApplication = Nothing
            excelWorkbook = Nothing
        End Try

        ' You can use the following method to automatically open the PDF after export if you wish
        ' Make sure that the file actually exists first...
        If (System.IO.File.Exists(outputPath)) Then
            'System.Diagnostics.Process.Start(outputPath)
        End If

        Return exportSuccessful
    End Function

    Public Sub ExportExcelToPDF(ByVal sourceFilePath As String, ByVal destinationFilePath As String)

        Dim myExcelApp As Microsoft.Office.Interop.Excel.Application
        Dim myExcelWorkbooks As Microsoft.Office.Interop.Excel.Workbooks = Nothing
        Dim myExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing


        Try

            Dim misValue As Object = System.Reflection.Missing.Value
            myExcelApp = New Microsoft.Office.Interop.Excel.ApplicationClass()
            myExcelApp.Visible = False
            Dim varMissing As Object = Type.Missing
            myExcelWorkbooks = myExcelApp.Workbooks

            'if file already exist then delete the file
            If (System.IO.File.Exists(destinationFilePath)) Then
                System.IO.File.Delete(destinationFilePath)
            End If
            myExcelWorkbook = myExcelWorkbooks.Open(sourceFilePath, misValue, misValue, _
                                                    misValue, misValue, misValue, misValue, misValue, misValue, _
                                                    misValue, misValue, misValue, misValue, misValue, misValue)
            myExcelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, _
                                                destinationFilePath, _
                                                Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, _
                                                varMissing, False, varMissing, varMissing, False, varMissing)
            myExcelWorkbooks.Close()
            myExcelApp.Quit()
        Catch ex As Exception

            Console.WriteLine("exception")
        Finally

            myExcelApp = Nothing
        End Try
    End Sub
End Class
