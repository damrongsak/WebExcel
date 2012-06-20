#Region ".NET Framework Class Import"
Imports System.Security
Imports System.Security.Principal
Imports System.Threading.Thread
Imports System.Net.Mail
Imports System.Data
#End Region



'Public Module Report

'    Public Function GetReportData(ByVal ReportName As String, ByVal ParamData As String, Optional ByVal Format As String = "PDF", Optional ByVal KeepFile As String = "", Optional ByVal OriginalFile As String = "", Optional ByVal FolderFile As String = "") As Byte()
'        Dim ShowFile As String = ""
'        Dim FS As Byte() = Nothing

'        Try
'            ShowFile = ExportReport(ReportName, ParamData, Format, OriginalFile, FolderFile)
'            If ShowFile <> "" Then FS = ReadFile(ShowFile)

'            Return FS
'        Catch ex As Exception
'            Throw ex
'        Finally
'            If ShowFile <> "" AndAlso KeepFile = "" Then DeleteFile(ShowFile)
'        End Try
'    End Function

'    Private Function ReadFile(ByVal FileName As String) As Byte()
'        Dim fs As System.IO.FileStream = Nothing
'        Dim lngLen As Long

'        Try ' Read file and return contents
'            fs = System.IO.File.Open(FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
'            lngLen = fs.Length

'            Dim Buffer(CInt(lngLen - 1)) As Byte
'            fs.Read(Buffer, 0, CInt(lngLen))

'            Return Buffer
'        Catch ex As Exception
'            Throw ex
'        Finally
'            If Not IsNothing(fs) Then
'                fs.Close()
'                fs = Nothing
'            End If
'        End Try
'    End Function

'    Public Function ExportReport(ByVal ReportName As String, ByVal ParamData As String, ByVal Format As String, Optional ByVal OriginalFile As String = "", Optional ByVal FolderFile As String = "") As String
'        Dim rpe As New ReportExecution.ReportExecutionService
'        Dim ei As New ReportExecution.ExecutionInfo
'        Dim Params() As ReportExecution.ParameterValue = Nothing
'        Dim Warning As ReportExecution.Warning() = Nothing

'        Dim RptName As String
'        Dim Results As Byte() = Nothing
'        Dim StreamID As String() = Nothing
'        Dim DeviceInfo As String = Nothing
'        Dim Encoding As String = String.Empty
'        Dim MimeType As String = String.Empty
'        Dim Extension As String = String.Empty
'        Dim HistoryID As String = Nothing

'        Dim GenFileName As String = ""

'        Try
'            'Set Credential
'            If ServerAuthenUsername = "" Then
'                rpe.Credentials = System.Net.CredentialCache.DefaultCredentials
'            Else
'                rpe.Credentials = New System.Net.NetworkCredential(ServerAuthenUsername, ServerAuthenPassword, ServerAuthenDomain)
'            End If

'            'Init Value
'            RptName = ReportServiceName & ReportName
'            Params = GenParams(ParamData)

'            'Render
'            ei = rpe.LoadReport(RptName, HistoryID)
'            rpe.SetExecutionParameters(Params, "en-us")
'            Results = rpe.Render(Format, DeviceInfo, Extension, Encoding, MimeType, Warning, StreamID)
'            If Not IsNothing(Results) AndAlso Results.Length > 0 Then
'                If FolderFile <> "" AndAlso OriginalFile <> "" Then
'                    GenFileName = HttpContext.Current.Server.MapPath(gFilePath) & FolderFile & OriginalFile
'                Else
'                    GenFileName = HttpContext.Current.Server.MapPath(ReportPath & Rnd() & "." & Format)
'                End If
'                System.IO.File.WriteAllBytes(GenFileName, Results)
'            End If

'            Return GenFileName
'        Catch ex As Exception
'            Throw ex
'        End Try
'    End Function

'    Public Function GenParams(ByVal ParamData As String) As ReportExecution.ParameterValue()
'        Dim ParamArr() As String = Nothing
'        Dim ParamVal As String
'        Dim i, j As Integer

'        Try
'            If ParamData <> "" Then ParamArr = ParamData.Split(",")
'            If Not IsNothing(ParamArr) AndAlso ParamArr.Length > 1 Then
'                'Dim Params(4) As ReportExecution.ParameterValue

'                'Params(0) = New ReportExecution.ParameterValue
'                'Params(0).Name = "FromDate"
'                'Params(0).Value = "03/01/2009"

'                'Params(1) = New ReportExecution.ParameterValue
'                'Params(1).Name = "ToDate"
'                'Params(1).Value = "06/01/2009"

'                'Params(2) = New ReportExecution.ParameterValue
'                'Params(2).Name = "BranchID"
'                'Params(2).Value = "01"

'                'Params(3) = New ReportExecution.ParameterValue
'                'Params(3).Name = "BranchID"
'                'Params(3).Value = "02"

'                Dim Params((ParamArr.Length / 2) - 1) As ReportExecution.ParameterValue
'                i = 0 : j = 0
'                Do While i < ParamArr.Length
'                    For Each ParamVal In ParamArr(i + 1).Split("|")
'                        If ParamVal.Trim <> "" Then
'                            If j + 1 >= Params.Length Then
'                                ReDim Preserve Params(j)
'                            End If
'                            Params(j) = New ReportExecution.ParameterValue
'                            Params(j).Name = ParamArr(i)
'                            Params(j).Value = ParamVal
'                            j += 1
'                        End If
'                    Next

'                    i += 2
'                Loop

'                Return Params
'            Else
'                Return Nothing
'            End If
'        Catch ex As Exception
'            Throw ex
'        End Try
'    End Function

'End Module


