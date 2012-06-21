
Partial Public Class ExcelToPDF
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        'Dim exc As ExcelToPDF = New ExcelToPDF()
        'exc.ToPDF()
        Dim filePath As String = Server.MapPath("Excels")
        Dim d As String = DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss")
        Dim srcFile As String = filePath + "\sample.xlsx"
        Dim descFile As String = filePath + "\" + d + "_test.pdf"
        Dim exc As MsExcelToPDF = New MsExcelToPDF()
        exc.ExportWorkbookToPdf(srcFile, descFile)
        'exc.ExportExcelToPDF(srcFile, descFile)
    End Sub
End Class