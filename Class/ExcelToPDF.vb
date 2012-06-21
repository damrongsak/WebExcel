'extern alias SpirePdf;
Imports Spire.Xls
Imports Spire.Xls.Converter
Imports Spire.Pdf
Imports pdf = Spire.Pdf
'Imports Excel = Microsoft.Office.Interop.Excel
'http://everlasting129.weebly.com/1/post/2012/04/convert-excel-to-pdf-with-c-vbnet.html
Public Class ExcelToPDF

    Public Sub ToPDF()
        Dim workbook As New Workbook()
        Dim filePath As String = Server.MapPath("Excels")
        Dim xls As String = filePath + "\sample.xlsx"
        workbook.LoadFromFile(xls, ExcelVersion.Version2010)

        Dim pdfConverter As New PdfConverter(workbook)
        Dim pdfDocument As New PdfDocument()
        pdfDocument.PageSettings.Orientation = pdf.PdfPageOrientation.Landscape
        pdfDocument.PageSettings.Width = 970
        pdfDocument.PageSettings.Height = 850

        Dim settings As New PdfConverterSettings()
        settings.TemplateDocument = pdfDocument
        pdfDocument = pdfConverter.Convert(settings)
        Dim pdffile As String = filePath + "\test.pdf"
        pdfDocument.SaveToFile(pdffile)
    End Sub

End Class

