Imports Microsoft.Office.Interop
Imports System.Security.Principal
Imports System.Globalization
Imports System.IO

Partial Public Class _Default
    Inherits System.Web.UI.Page

    ' if error : http://greatfriends.biz/webboards/msg.asp?id=126431
    ' chart : http://support.microsoft.com/kb/219151
    ' http://www.vbdotnetheaven.com/uploadfile/ggaganesh/excelspreadsheet04182005093012am/excelspreadsheet.aspx
    ' http://msdn.microsoft.com/en-us/library/aa188489(office.10).aspx
    ' http://www.expert2you.com/view_article.php?art_id=3265
    ' http://en.wikipedia.org/wiki/Visual_Studio_Tools_for_Office
    ' conf : http://forums.asp.net/t/1303594.aspx/1
    '
    '
    'Keep the application object and the workbook object global, so you can  
    'retrieve the data in Button2_Click that was set in Button1_Click.
    Dim objApp As Excel.Application
    Dim objBook As Excel._Workbook

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim curentthread As System.Threading.Thread
        curentthread = System.Threading.Thread.CurrentThread
        curentthread.CurrentCulture = New CultureInfo("en-US")
        Dim cname As String = WindowsIdentity.GetCurrent().Name
        Dim name As String = User.Identity.Name
        Response.Write("WindowsIdentity.GetCurrent().Name : " & cname & ", User.Identity.Name : " & name)

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        'Dim workbookPath As String = Server.MapPath("~") & "ngv.xlsx"
        'Dim excelApp = New Excel.Application()
        'Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Open(workbookPath, 0, False, 5, _
        '"", "", False, Excel.XlPlatform.xlWindows, "", True, False, 0, True)
        ''Return control of Excel to the user.
        'excelApp.Visible = True
        'excelApp.UserControl = True

        ''Clean up a little.
        'excelWorkbook = Nothing

        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range

        ' Create a new instance of Excel and start a new workbook.
        objApp = New Excel.Application()
        objBooks = objApp.Workbooks
        objBook = objBooks.Add
        objSheets = objBook.Worksheets
        objSheet = objSheets(1)

        'Get the range where the starting cell has the address
        'm_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
        range = objSheet.Range("A1", Reflection.Missing.Value)
        range = range.Resize(5, 5)

        If (Me.FillWithStrings.Checked = False) Then
            'Create an array.
            Dim saRet(5, 5) As Double

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5

                    'Put a counter in the cell.
                    saRet(iRow, iCol) = iRow * iCol
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet

        Else
            'Create an array.
            Dim saRet(5, 5) As String

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5

                    'Put the row and column address in the cell.
                    saRet(iRow, iCol) = iRow.ToString() + "|" + iCol.ToString()
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet
        End If

        Dim d As String = DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss")
        Dim xlsx As String = String.Format("Report_{0}.xlsx", d)
        Dim filePath As String = String.Format("{0}\{1}", Server.MapPath("Excels"), xlsx)

        objSheet.SaveAs(filePath)
        'Return control of Excel to the user.
        'objApp.Visible = True
        'objApp.UserControl = True

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
        objBooks = Nothing
        Dim exc As Excel_2007 = New Excel_2007()
        exc.Kill("EXCEL")

        Response.Write(String.Format("<br /><a href='Excels/{0}' >downlaod: {0}</a>", xlsx))
        'Dim stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)
        'Response.Clear()
        'Response.AddHeader("Content-Disposition", "attachment; filename=" + xlsx)
        'Response.AddHeader("Content-Length", stream.Length.ToString())
        'Response.ContentType = "application/octet-stream"
        'stream.Close()
        'Response.WriteFile("Excels/" & xlsx)
        'Response.End()

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range

        'Get a reference to the first sheet of the workbook.
        On Error GoTo ExcelNotRunning
        objSheets = objBook.Worksheets
        objSheet = objSheets(1)

ExcelNotRunning:
        If (Not (Err.Number = 0)) Then
            Response.Write("Cannot find the Excel workbook.  Try clicking Button1 to " + _
            "create an Excel workbook with data before running Button2." + _
            "Missing Workbook?")

            'We cannot automate Excel if we cannot find the data we created, 
            'so leave the subroutine.
            Exit Sub
        End If

        'Get a range of data.
        range = objSheet.Range("A1", "E5")

        'Retrieve the data from the range.
        Dim saRet(,) As Object
        saRet = range.Value

        'Determine the dimensions of the array.
        Dim iRows As Long
        Dim iCols As Long
        iRows = saRet.GetUpperBound(0)
        iCols = saRet.GetUpperBound(1)

        'Build a string that contains the data of the array.
        Dim valueString As String
        valueString = "Array Data" + vbCrLf

        Dim rowCounter As Long
        Dim colCounter As Long
        For rowCounter = 1 To iRows
            For colCounter = 1 To iCols

                'Write the next value into the string.
                valueString = String.Concat(valueString, _
                    saRet(rowCounter, colCounter).ToString() + ", ")

            Next colCounter

            'Write in a new line.
            valueString = String.Concat(valueString, vbCrLf)
        Next rowCounter

        'Report the value of the array.
        Response.Write(valueString + "Array Values")

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
        Dim exc As Excel_2007 = New Excel_2007()
        exc.Kill("EXCEL")

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        
        'http://exceldatareader.codeplex.com/
        'http://www.thaicreate.com/dotnet/csharp-dot-net-generate-excel.html
        'http://greatfriends.biz/webboards/msg.asp?id=80114

        'http://support.microsoft.com/kb/302084
        'http://72.15.199.198/articles/Creating_Spreadsheets_Server.aspx
        'http://www.codeproject.com/KB/cs/Write_Data_to_Excel_using.aspx
        'http://www.c-sharpcorner.com/UploadFile/ggaganesh/CreateExcelSheet12012005015333AM/CreateExcelSheet.aspx
        'http://www.aspnetpro.com/NewsletterArticle/2003/09/asp200309so_l/asp200309so_l.asp
        'http://www.codeguru.com/csharp/.net/net_asp/tutorials/article.php/c13123/

        Dim d As String = DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss")
        Dim xlsx As String = String.Format("Report_{0}.xlsx", d)
        Dim filePath As String = String.Format("{0}\{1}", Server.MapPath("Excels"), xlsx)

        Dim exc As Excel_2007 = New Excel_2007()
        exc.filePath = filePath
        exc.Run()

        Dim stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment; filename=" + xlsx)
        Response.AddHeader("Content-Length", stream.Length.ToString())
        Response.ContentType = "application/octet-stream"
        stream.Close()
        Response.WriteFile("Excels/" & xlsx)
        Response.End()

    End Sub
End Class