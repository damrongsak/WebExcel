#Region ".NET base class name space imports"
Imports System.Configuration
Imports System.Web
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
#End Region

Public Class CRReportComponent
    Private _DB_Provider As String
    Private _DB_UserName As String
    Private _DB_Password As String
    Private _DB_DataSource As String
    Private _DB_Name As String

    Dim ReportFilePath, FilePath As String

    Public Sub New()
        ReadDALConfigurations()
    End Sub

    Private Sub ReadDALConfigurations()
        Dim Encrypt As New SecurityUtil

        Try
            _DB_Provider = ConfigurationManager.AppSettings("DB_Provider") & ""
            _DB_DataSource = ConfigurationManager.AppSettings("DB_DataSource") & ""
            _DB_Name = ConfigurationManager.AppSettings("DB_Name") & ""
            _DB_UserName = Encrypt.DecryptData(ConfigurationManager.AppSettings("DB_UserName") & "") & ""
            _DB_Password = Encrypt.DecryptData(ConfigurationManager.AppSettings("DB_Password") & "") & ""

            ReportFilePath = ConfigurationManager.AppSettings("ReportFilePath") & ""
            FilePath = ConfigurationManager.AppSettings("FilePath") & ""
        Catch ex As Exception
            ClearObject(Encrypt)
        End Try
    End Sub

    Public Function GetReportData(ByVal FileName As String, ByVal ParamData As String, Optional ByVal Format As String = "pdf", Optional ByVal FSelection As String = "") As Byte()

        Dim ShowFile As String = ""
        Dim FS As Byte() = Nothing

        Try
            ShowFile = ExportReport(FileName, ParamData, Format, FSelection) & ""
            If ShowFile <> "" Then FS = ReadFile(ShowFile)

            Return FS
        Catch ex As Exception
            Throw New DALException(ex.Message)
        Finally
            If File.Exists(ShowFile) = True Then
                DeleteFile(ShowFile)
            End If
            'If Directory.Exists(ShowFile) Then DeleteFile(ShowFile)
        End Try
    End Function

    Private Function ReadFile(ByVal FileName As String) As Byte()
        Dim fs As FileStream = Nothing
        Dim lngLen As Long

        Try
            ' Read file and return contents
            fs = File.Open(FileName, FileMode.Open, FileAccess.Read)
            lngLen = fs.Length

            Dim Buffer(CInt(lngLen - 1)) As Byte
            fs.Read(Buffer, 0, CInt(lngLen))

            Return Buffer
        Catch ex As Exception
            Throw New DALException("Unable to read file : " & ex.Message)
        Finally
            If Not IsNothing(fs) Then
                fs.Close()
                fs.Dispose()
                fs = Nothing
            End If
        End Try
    End Function


    Private Sub SetReportDB(ByRef rptDoc As ReportDocument)
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo

        For Each tbCurrent In rptDoc.Database.Tables
            tliCurrent = tbCurrent.LogOnInfo
            '    With tliCurrent.ConnectionInfo
            '        .ServerName = _DB_DataSource
            '        .UserID = _DB_UserName
            '        .Password = _DB_Password
            '        .DatabaseName = _DB_Name
            'End With
            With tliCurrent.ConnectionInfo
                .ServerName = "XE"
                .UserID = "tracking"
                .Password = "P@ssw0rd1"
                .DatabaseName = "XE"
            End With
            tbCurrent.ApplyLogOnInfo(tliCurrent)
        Next tbCurrent
    End Sub

    Private Sub InitReport(ByRef oRptDoc As ReportDocument, ByVal FileName As String, ByVal ParamData As String, ByVal FSelection As String, Optional ByVal Conn As String = "")
        'Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        'Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim PData() As String
        Dim I As Integer = 0
        Dim Sect As Section
        Dim srpt As ReportDocument
        Dim robj As ReportObject
        Dim filepath As String

        Try
            filepath = HttpContext.Current.Server.MapPath(ReportFilePath) & FileName
            oRptDoc.Load(filepath)
            Dim tm As New Globalization.CultureInfo("th-TH")
            If ParamData <> "" Then
                PData = Split(ParamData, ",")
                Do While I < UBound(PData)
                    oRptDoc.SetParameterValue(PData(I), PData(I + 1).ToString(tm))
                    I += 2
                Loop
            End If

            If FSelection & "" <> "" Then
                'Replace ส่วนที่เป็น Like
                FSelection = Replace(FSelection, "/*", "'*")
                FSelection = Replace(FSelection, "*/", "*'")
                FSelection = Replace(FSelection, "/", "'")
                oRptDoc.RecordSelectionFormula = FSelection
            End If

            SetReportDB(oRptDoc)

            For Each Sect In oRptDoc.ReportDefinition.Sections
                For Each robj In Sect.ReportObjects
                    If robj.Kind = ReportObjectKind.SubreportObject Then
                        srpt = oRptDoc.OpenSubreport(CType(robj, SubreportObject).SubreportName)
                        SetReportDB(srpt)
                    End If
                Next
            Next

        Catch ex As Exception
            Throw New DALException("Unable to load report : " & ex.ToString)
        End Try
    End Sub

    Public Function ExportReport(ByVal FileName As String, ByVal ParamData As String, Optional ByVal Format As String = "pdf", Optional ByVal FSelection As String = "") As String
        Dim ExportPath As String
        Dim SaveName As String
        Dim crExportOptions As ExportOptions
        Dim crDiskFileDestinationOptions As DiskFileDestinationOptions
        Dim FName As String
        Dim Export_File As String
        Dim oRptDoc As New ReportDocument

        Dim Obj As Object = Nothing
        Dim Conn As String = ""

        Try
            If InStr(FileName, "|") > 0 Then
                Obj = FileName.Split("|")
                If Not IsNothing(Obj) AndAlso Obj.Length > 0 Then
                    FileName = Obj(0)
                    Conn = Obj(1)
                End If
            End If

            InitReport(oRptDoc, FileName, ParamData, FSelection, Conn)

            FName = Replace(FileName.Substring(FileName.LastIndexOf("/") + 1), ".rpt", "")
            ExportPath = HttpContext.Current.Server.MapPath(FilePath)
            If Directory.Exists(ExportPath) = False Then
                Directory.CreateDirectory(ExportPath)
            End If

            crDiskFileDestinationOptions = New DiskFileDestinationOptions
            crExportOptions = oRptDoc.ExportOptions
            crExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            SaveName = FName

            Randomize()
            SaveName = SaveName & "_" & CStr(Rnd())

            With crExportOptions
                Select Case Format.ToUpper()
                    Case "PDF", "PORTABLE DOCUMENT"
                        Format = "PDF"
                        .DestinationOptions = crDiskFileDestinationOptions
                        .ExportFormatType = ExportFormatType.PortableDocFormat
                    Case "XLS", "MS EXCEL (XLS)"
                        Format = "XLS"
                        .ExportFormatType = ExportFormatType.Excel
                        .DestinationOptions = crDiskFileDestinationOptions
                    Case "DOC", "MS WORD (DOC)"
                        Format = "DOC"
                        .ExportFormatType = ExportFormatType.WordForWindows
                        .DestinationOptions = crDiskFileDestinationOptions
                    Case "RTF", "RICH TEXT FILE"
                        Format = "RTF"
                        .ExportFormatType = ExportFormatType.RichText
                        .DestinationOptions = crDiskFileDestinationOptions
                End Select
            End With

            SaveName &= "." & Format
            crDiskFileDestinationOptions.DiskFileName = ExportPath & SaveName
            Export_File = ExportPath & SaveName
            'Export_File = "../files/Reports/" & SaveName
            oRptDoc.Export()

            Return Export_File
        Catch ex As Exception
            Throw New DALException("Unable to get report file : " & ex.ToString)
        End Try
    End Function

    'Public Function SaveReportFile(ByVal FileName As String, ByVal ParamData As String, Optional ByVal Format As String = "pdf", Optional ByVal OwnFileName As String = "", Optional ByVal FSelection As String = "", Optional ByVal RFolder As String = "") As String
    '    Dim ExportPath As String
    '    Dim SaveName As String
    '    Dim crExportOptions As ExportOptions
    '    Dim crDiskFileDestinationOptions As DiskFileDestinationOptions
    '    Dim FName As String
    '    Dim Export_File As String
    '    Dim oRptDoc As New ReportDocument

    '    Try
    '        Try
    '            InitReport(oRptDoc, FileName, ParamData, FSelection)
    '        Catch ex As Exception
    '            Throw New DALException("Unable to Init file : " & ex.ToString)
    '        End Try

    '        FName = Replace(FileName.Substring(FileName.LastIndexOf("/") + 1), ".rpt", "")
    '        ExportPath = HttpContext.Current.Server.MapPath(FilePath)
    '        If Directory.Exists(ExportPath) = False Then
    '            Directory.CreateDirectory(ExportPath)
    '        End If

    '        crDiskFileDestinationOptions = New DiskFileDestinationOptions
    '        crExportOptions = oRptDoc.ExportOptions
    '        crExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
    '        SaveName = FName

    '        Randomize()
    '        SaveName = SaveName & "_" & CStr(Rnd())
    '        If OwnFileName <> "" Then
    '            SaveName = OwnFileName
    '        End If

    '        With crExportOptions
    '            Select Case Format.ToUpper()
    '                Case "PDF", "PORTABLE DOCUMENT"
    '                    Format = "PDF"
    '                    .DestinationOptions = crDiskFileDestinationOptions
    '                    .ExportFormatType = ExportFormatType.PortableDocFormat
    '                Case "XLS", "MS EXCEL (XLS)"
    '                    Format = "XLS"
    '                    .ExportFormatType = ExportFormatType.Excel
    '                    .DestinationOptions = crDiskFileDestinationOptions
    '                Case "DOC", "MS WORD (DOC)"
    '                    Format = "DOC"
    '                    .ExportFormatType = ExportFormatType.WordForWindows
    '                    .DestinationOptions = crDiskFileDestinationOptions
    '                Case "RTF", "RICH TEXT FILE"
    '                    Format = "RTF"
    '                    .ExportFormatType = ExportFormatType.RichText
    '                    .DestinationOptions = crDiskFileDestinationOptions
    '            End Select
    '        End With

    '        SaveName &= "." & Format

    '        If RFolder <> "" Then
    '            crDiskFileDestinationOptions.DiskFileName = ExportPath & RFolder & SaveName
    '            Export_File = ExportPath & RFolder & SaveName
    '        Else
    '            crDiskFileDestinationOptions.DiskFileName = ExportPath & SaveName
    '            Export_File = ExportPath & SaveName
    '        End If

    '        'Export_File = "../files/Reports/" & SaveName
    '        oRptDoc.Export()
    '        'Return Export_File
    '        Return ""
    '    Catch ex As Exception
    '        Throw New DALException("Unable to save report file : " & ex.ToString)
    '    End Try
    'End Function

    Public Function ExportReport2(ByVal FileName As String, ByVal ParamData As String, Optional ByVal Format As String = "pdf", Optional ByVal FSelection As String = "", Optional ByVal ReportNameDesc As String = "", Optional ByVal tmpPath As String = "") As String
        Dim ExportPath As String
        Dim SaveName As String
        Dim crExportOptions As ExportOptions
        Dim crDiskFileDestinationOptions As DiskFileDestinationOptions
        Dim FName As String
        Dim Export_File As String
        Dim oRptDoc As New ReportDocument

        Dim Obj As Object = Nothing
        Dim Conn As String = ""

        Try
            If InStr(FileName, "|") > 0 Then
                Obj = FileName.Split("|")
                If Not IsNothing(Obj) AndAlso Obj.Length > 0 Then
                    FileName = Obj(0)
                    Conn = Obj(1)
                End If
            End If

            InitReport(oRptDoc, FileName, ParamData, FSelection, Conn)

            FName = Replace(FileName.Substring(FileName.LastIndexOf("/") + 1), ".rpt", "")
            'ExportPath = HttpContext.Current.Server.MapPath(FilePath)



            crDiskFileDestinationOptions = New DiskFileDestinationOptions
            crExportOptions = oRptDoc.ExportOptions
            crExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            'ReportNameDesc = FormatRetrieveNameFile(ReportNameDesc)
            SaveName = FName & ReportNameDesc
            'Randomize()
            'SaveName = SaveName & "_" & CStr(Rnd())

            With crExportOptions
                Select Case Format.ToUpper()
                    Case "PDF", "PORTABLE DOCUMENT"
                        Format = "PDF"
                        .DestinationOptions = crDiskFileDestinationOptions
                        .ExportFormatType = ExportFormatType.PortableDocFormat
                    Case "XLS", "MS EXCEL (XLS)"
                        Format = "XLS"
                        .ExportFormatType = ExportFormatType.Excel
                        .DestinationOptions = crDiskFileDestinationOptions
                    Case "DOC", "MS WORD (DOC)"
                        Format = "DOC"
                        .ExportFormatType = ExportFormatType.WordForWindows
                        .DestinationOptions = crDiskFileDestinationOptions
                    Case "RTF", "RICH TEXT FILE"
                        Format = "RTF"
                        .ExportFormatType = ExportFormatType.RichText
                        .DestinationOptions = crDiskFileDestinationOptions
                End Select
            End With

            SaveName &= "." & Format
            ExportPath = HttpContext.Current.Server.MapPath(gFilePath & tmpPath)
            'If Directory.Exists(ExportPath) = False Then
            '    Directory.CreateDirectory(ExportPath)
            'End If
            crDiskFileDestinationOptions.DiskFileName = ExportPath & SaveName
            Export_File = ExportPath & SaveName
            'Export_File = "../files/Daily/" & SaveName
            If File.Exists(Export_File) = True Then
                DeleteFile(Export_File)
            End If
            oRptDoc.Export()
            'oRptDoc.Dispose()
            Return Export_File
        Catch ex As Exception
            Throw New DALException("Unable to get report file : " & ex.ToString)
        End Try
    End Function
End Class
