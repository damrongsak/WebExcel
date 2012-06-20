#Region ".NET Framework Class Import"
Imports System.Security
Imports System.Security.Principal
Imports System.Threading.Thread
Imports System.Net.Mail
Imports System.Data
#End Region



Public Module Project

    Public DAL As New DALComponent
    Public DB As New DBUTIL()
    'Public DAL2 As New DALComponent("2")
    Public BLL As New BLLComponent
    Public RPT As New CRReportComponent

    Public Const cCookieExpiration As Integer = 10 ' Minutes

    Public Const opINSERT As Integer = 1
    Public Const opUPDATE As Integer = 2
    Public Const opDELETE As Integer = 3

    Public Const CellNumeric As Integer = 1
    Public Const CellText As Integer = 2
    Public Const CellDate As Integer = 3
    Public Const CellExpression As Integer = 4
    Public Const CellBoolean As Integer = 5
    Public Const CellDateTime As Integer = 6

    Public gReadConfig As String
    Public gDebugLevel As String
    Public gUseCookies As Boolean = True
    Public gCookiesName As String = "System"

    Public gSMTP_Server As String
    Public gSender_EMail As String
    Public gMailMode As String
    Public gTest_EMail As String
    Public gEmailDelimeter As String
    Public gEmail_ProgramName As String

    Public gFilePath As String


    'Public gVettingEMail As String

    '' ****** Log Category Constants ******
    'Public Const LogLogon As String = "LOGON"
    'Public Const LogError As String = "ERROR"
    'Public Const LogMail As String = "MAIL"
    'Public Const LogPurchaseSIRE As String = "PURCHASE SIRE"
    'Public Const LogDownloadInspectionReport As String = "DOWNLOAD INSPECTION"
    'Public Const LogViewInspectionReport As String = "VIEW INSPECTION"
    'Public Const LogRequest As String = "REQUEST"
    'Public Const LogScreening As String = "SCREENING"
    'Public Const LogClearanceSubmit As String = "SUBMIT CLEARANCE"
    'Public Const LogClearanceVerify As String = "VERIFY CLEARANCE"
    'Public Const LogClearanceFinal As String = "FINAL CLEARANCE"
    'Public Const LogRelease As String = "RELEASE"
    'Public Const LogAppealingSubmit As String = "SUBMIT APPEALING"
    'Public Const LogAppealingVerify As String = "VERIFY APPEALING"
    'Public Const LogAppealingFinal As String = "FINAL APPEALING"
    'Public Const LogCancel As String = "CANCEL"

    Public gFileType As String
    Public gResumePath As String
    Public NearExpInspectionReport As String

    Public AgeCriteria As Integer = 20 'Default
    Public Sinfo_Expired As Integer = 1 'Default

    ' ****** Log Category Constants ******
    Public Const catAppLog As String = "LOG"
    Public Const catAddLog As String = "Add"
    Public Const catUpdateLog As String = "Update"
    Public Const catDeleteLog As String = "Delete"
    Public Const catErrorLog As String = "ERROR"
    Public Const catMailLog As String = "MAIL"
    Public Const catViewLog As String = "VIEW"
    Public Const catPrintLog As String = "PRINT"
    Public Const catBannedLog As String = "BANNED"
    Public Const catBannedRequestLog As String = "BANNED_REQUEST"


    ' ******* FCK Editor ********
    Public ProjectName, ProjectServer, ToolbarSet, BasePath, EditorUploadPath, ImgPath, LinkPath As String

    ' ******* Upload File Type ********
    Public imgFileType, docFileType, clipFileType, soundFileType, flashFileType, aFileType As String

    ' ******* Reporting Service ********
    Public ServerAuthenDomain, ServerAuthenUsername, ServerAuthenPassword As String
    Public ReportServiceName, ReportPath As String

    ' ******* URL ********
    Public URL_Main As String

    ' ******* Axis Camera ********* 'pui 13/5/52    
    Public AutoStart, UIMode, MediaURL, MediaUsername, MediaPassword, MediaType, CameraPath As String

    Public SendMailEncryptURL As String

    Public gBannedDuration As String

#Region "Initial Config Value"
    Public Sub ReadConfigurations()
        gReadConfig = "Y"

        With HttpContext.Current
            gDebugLevel = ConfigurationManager.AppSettings("DebugLevel") & ""
            gUseCookies = (ConfigurationManager.AppSettings("UseCookies") & "").ToLower() = "true"

            gSMTP_Server = ConfigurationManager.AppSettings("SMTP_Server") & ""
            gSender_EMail = ConfigurationManager.AppSettings("Sender_EMail") & ""
            gTest_EMail = ConfigurationManager.AppSettings("TEST_EMail") & ""
            gEmail_ProgramName = ConfigurationManager.AppSettings("Sender_Name") & ""
            gMailMode = ConfigurationManager.AppSettings("MailMode") & ""


            gFileType = ConfigurationManager.AppSettings("FileType") & ""
            gFilePath = ConfigurationManager.AppSettings("FilePath") & ""



            '// FCK Editor
            ProjectName = ConfigurationManager.AppSettings("ProjectName") & ""
            ProjectServer = ConfigurationManager.AppSettings("ProjectServer") & ""
            'ToolbarSet = ConfigurationManager.AppSettings("ToolbarSet") & ""
            'BasePath = ConfigurationManager.AppSettings("BasePath") & ""
            'EditorUploadPath = ConfigurationManager.AppSettings("EditorUploadPath") & ""
            ImgPath = ConfigurationManager.AppSettings("ImgPath") & ""
            'LinkPath = ConfigurationManager.AppSettings("LinkPath") & ""


            '// Reporting Services
            'ServerAuthenDomain = ConfigurationManager.AppSettings("ServerAuthenDomain") & ""
            'ServerAuthenUsername = ConfigurationManager.AppSettings("ServerAuthenUsername") & ""
            'ServerAuthenPassword = ConfigurationManager.AppSettings("ServerAuthenPassword") & ""
            'ReportServiceName = ConfigurationManager.AppSettings("ReportServiceName") & ""
            ReportPath = ConfigurationManager.AppSettings("ReportPath") & ""

            '// Upload File Type
            imgFileType = ConfigurationManager.AppSettings("imgType") & ""
            docFileType = ConfigurationManager.AppSettings("docType") & ""
            clipFileType = ConfigurationManager.AppSettings("clipType") & ""
            soundFileType = ConfigurationManager.AppSettings("soundType") & ""
            flashFileType = ConfigurationManager.AppSettings("flashType") & ""
            aFileType = ConfigurationManager.AppSettings("aFileType") & ""
            '// URL
            'URL_Main = ConfigurationManager.AppSettings("MAIN_URL") & ""



            'SendMailEncryptURL = ConfigurationManager.AppSettings("SendMailEncryptURL") & ""

            gBannedDuration = ConfigurationManager.AppSettings("BannedDuration") & ""
        End With
    End Sub
#End Region

#Region "EMail Management"
    'Public Function SendEMail(ByVal Subject As String, ByVal Sender As String, ByVal RecpName As String, ByVal RecpCompany As String, ByVal RecpNo As String, ByVal Message As String, ByVal Filename() As String, ByVal Priority As Integer, ByVal BillingCode As String) As String
    '    Dim objMail As New MailMessage
    '    Dim I As Integer
    '    Dim ErrMsg As String
    '    Dim SendMsg As String

    '    Try
    '        'objMail.To = RecpNo
    '        'objMail.From = Sender

    '        objMail = New MailMessage(Sender, RecpNo)

    '        objMail.Subject = Subject
    '        SendMsg = Subject + "<br><br>" + "รายชื่อเอกสารที่แนบมา <br>"


    '        If Not IsNothing(Filename) AndAlso Filename.Length > 0 Then
    '            For I = 0 To UBound(Filename)
    '                SendMsg += GetFileName(Filename(I)) + "<br>"
    '                objMail.Attachments.Add(New Attachment(Filename(I)))
    '            Next I
    '        End If

    '        objMail.Body = SendMsg + "<br> หมายเหตุ : " + Message
    '        objMail.IsBodyHtml = True


    '        Dim SMTPMail As New SmtpClient

    '        SMTPMail.Host = gSMTP_Server
    '        SMTPMail.Send(objMail)

    '        ErrMsg = ""
    '    Catch ex As Exception
    '        ErrMsg = ex.Message
    '    End Try
    '    objMail = Nothing
    '    Return ErrMsg
    'End Function

    Public Function SendEMailData(ByVal Subject As String, ByVal Message As String _
        , ByVal Sender As String, ByVal Receiver As String, Optional ByVal MailCC As String = "" _
        , Optional ByVal Filename() As String = Nothing, Optional ByVal MailBCC As String = "") As String
        Dim SMTPMail As SmtpClient = Nothing
        Dim objMail As MailMessage = Nothing
        Dim CcMail As MailAddress = Nothing
        Dim BccMail As MailAddress = Nothing
        Dim ToMail As MailAddress = Nothing
        Dim FromMail As MailAddress = Nothing

        Dim MailFrom, MailTo As String
        Dim ErrMsg As String = ""
        Dim OneMail() As String
        Dim OneMailTo() As String

        Dim i As Integer

        Try
            If gMailMode = "3" Then 'Config Send Mail
                MailFrom = IIf(Sender = "", gSender_EMail, Sender)
                MailTo = IIf(gTest_EMail = "", Receiver, gTest_EMail)

                objMail = New MailMessage
                'If gEmail_ProgramName & "" <> "" Then
                '    FromMail = New MailAddress(MailFrom, gEmail_ProgramName)
                'Else
                '    FromMail = New MailAddress(MailFrom)
                'End If
                If Sender <> "" Then
                    FromMail = New MailAddress(MailFrom, HttpContext.Current.Session("USER_DESC") & "")
                Else
                    FromMail = New MailAddress(MailFrom, gEmail_ProgramName)
                End If
                objMail.From = FromMail
                objMail.Subject = Subject
                objMail.Body = Message
                If MailTo <> "" Then
                    OneMailTo = Split(MailTo, ";")
                    For i = 0 To OneMailTo.Length - 1
                        Try
                            ToMail = New MailAddress(OneMailTo(i))
                            objMail.To.Add(ToMail)
                        Catch
                        End Try
                    Next
                End If
                If MailCC <> "" Then
                    OneMail = Split(MailCC, ";")
                    For i = 0 To OneMail.Length - 1
                        Try
                            CcMail = New MailAddress(OneMail(i))
                            objMail.CC.Add(CcMail)
                        Catch
                        End Try
                    Next
                End If

                'BCC To Sender
                If MailBCC <> "" Then
                    OneMail = Split(MailBCC, ";")
                    For i = 0 To OneMail.Length - 1
                        Try
                            BccMail = New MailAddress(OneMail(i))
                            objMail.Bcc.Add(BccMail)
                        Catch
                        End Try
                    Next
                End If
                'objMail.Bcc.Add(FromMail)

                If Not IsNothing(Filename) AndAlso Filename.Length > 0 Then
                    For i = 0 To UBound(Filename)
                        Try
                            If Filename(i) & "" <> "" Then
                                If InStr(Filename(i), ":") > 0 Then 'มี Path แล้ว
                                    objMail.Attachments.Add(New Attachment((Filename(i))))
                                Else
                                    objMail.Attachments.Add(New Attachment(HttpContext.Current.Server.MapPath(Filename(i))))
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Next i
                End If

                objMail.IsBodyHtml = True
                SMTPMail = New SmtpClient()
                SMTPMail.Host = gSMTP_Server
                'SMTPMail.Port = 25
                SMTPMail.Send(objMail)

                System.Threading.Thread.Sleep(0)

                ErrMsg = ""
            End If
        Catch ex As Exception
            ErrMsg = GetErrorMsg(ex)
        Finally
            ClearObject(SMTPMail)
            ClearObject(objMail)
        End Try

        Return ErrMsg
    End Function
#End Region


    Public Function UploadFile(ByRef FileUpload As FileUpload, ByVal FilePath As String, ByVal eFileType As String, ByRef FileName As String, Optional ByRef OrgFileName As String = Nothing) As String
        Dim FileType, FullFileName As String
        Dim CanSave As Boolean = False
        Dim ret As String = ""

        Try
            'If FileUpload.HasFile AndAlso eFileType <> "" AndAlso FileName <> "" Then
            If FileUpload.HasFile Then
                If Not IsNothing(OrgFileName) Then OrgFileName = FileUpload.FileName
                FileType = (GetFileType(FileUpload.FileName) & "").ToLower()
                If eFileType = "" Then eFileType = imgFileType & "|" & docFileType 'Default
                If InStr("|" & eFileType & "|", "|" & FileType & "|") > 0 Then
                    CanSave = True
                End If

                If Not CanSave Then
                    ret = "FI" : FileName = ""
                Else
                    If FileName <> "" Then
                        FileName &= FileType
                    Else
                        FileName = FileUpload.FileName
                    End If
                    FullFileName = HttpContext.Current.Server.MapPath(FilePath & FileName)

                    FileUpload.SaveAs(FullFileName)
                End If
            Else
                FileName = ""
                ret = "File is not exist."
            End If

            Return ret
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetDateTimeString() As String
        Return Format(Now, "yyyyMMddHHmss")
    End Function

    Public Sub GetDefaultStart_EndDate(ByRef StartDate As String, ByRef EndDate As String)
        Dim D As Date = System.DateTime.Now

        Try
            StartDate = AppFormatDate(D.AddMonths(-3))
            EndDate = AppFormatDate(D)
        Catch ex As Exception
            StartDate = "" : EndDate = ""
        End Try
    End Sub

    Public Function CompareFactor(ByVal Oper As String, ByVal Value1 As Integer, ByVal Value2 As Integer, ByVal FactorValue As Integer) As Boolean
        Dim ret As Boolean = False

        Try
            Select Case Oper
                Case "EQ" : ret = IIf(FactorValue = Value1, True, False)
                Case "MO" : ret = IIf(FactorValue > Value1, True, False)
                Case "ME" : ret = IIf(FactorValue >= Value1, True, False)
                Case "BE" : ret = IIf(FactorValue >= Value1 AndAlso FactorValue <= Value2, True, False)
                Case "LO" : ret = IIf(FactorValue < Value1, True, False)
                Case "LE" : ret = IIf(FactorValue <= Value1, True, False)
            End Select

            Return ret
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub FillTable(ByRef T As Table, ByVal MaxRow As Integer, ByVal MaxCol As Integer, _
                         Optional ByVal CellValue As String = "&nbsp", Optional ByVal ShowBorder As Boolean = False, Optional ByVal FirstColWidth As Integer = -1)
        Dim R, C As Integer
        Dim TR As TableRow
        Dim TC As TableCell

        T = Nothing
        T = New Table
        For R = 0 To MaxRow - 1
            TR = New TableRow
            For C = 0 To MaxCol - 1
                TC = New TableCell
                TC.Text = CellValue
                If FirstColWidth > 0 Then TC.Width = Unit.Pixel(FirstColWidth)
                If ShowBorder Then
                    TC.BorderColor = Drawing.Color.Black
                    TC.BorderWidth = 1
                End If
                TR.Cells.Add(TC)
            Next
            T.Rows.Add(TR)
        Next
    End Sub


    'Updated By Aoy 04/05/2552
    Public Function LoadUserData(ByVal UserName As String, Optional ByVal Password As String = "", Optional ByVal From As String = "") As String
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim DR2 As DataRow = Nothing
        Dim Msg As String = ""
        Dim IDCard As String = ""
        Dim RoleDesc(1) As String
        Dim TerminalScope As String = ""
        Dim Criteria As String = ""
        Dim BranchList As String = ""
        'Dim RPTWS As RPT.ReportTo = Nothing

        Try
            '21/4/2551 in valid password because password start with > , < , ( , ).


            If UserName <> "" Then
                DT = DAL.Login(UserName, "")
                DR = GetDR(DT)
            End If

            If Not IsNothing(DR) Then
                If From <> "IsAuthenticated" AndAlso DR("PASSWORD") & "" <> Password Then
                    'If From <> "IsAuthenticated" AndAlso Password = "" Then
                    Msg = "Invalid user name or password."
                ElseIf DR("DISABLED_FLAG") & "" = "Y" Then
                    Msg = "User : " & UserName & " was banned! Please contact system administrator."
                Else
                    Msg = ""
                    With HttpContext.Current
                        RoleDesc(0) = DR("ROLE_ID") & ""
                        CreateContext(UserName, RoleDesc)
                        .Session("ROLES") = RoleDesc

                        .Session("USER_NAME") = DR("USER_NAME") & ""
                        .Session("USER_DESC") = Replace(DR("USER_DESC") & "", "*", "")
                        '.Session("EMP_CODE") = DR("EMP_CODE") & ""
                        .Session("ROLE_ID") = DR("ROLE_ID") & ""
                        .Session("ROLE_NAME") = DR("ROLE_DESC") & ""
                        .Session("USER_LEVEL") = DR("USER_LEVEL") & ""
                        .Session("USER_TYPE") = DR("USER_TYPE") & ""
                        .Session("GROUP_ID") = DR("GROUP_ID") & ""
                        .Session("GROUP_NAME") = DR("GROUP_NAME") & ""

                        .Session("RIGHTS") = DR("RIGHTS") & ""
                        .Session("PERMIS_INFOS") = DR("PERMIS_INFOS") & ""
                        .Session("PERMIS_SV_RESPONSE") = DR("PERMIS_SV_RESPONSE") & ""
                        .Session("PERMIS_PROJECT_TYPES") = DR("PERMIS_PROJECT_TYPES") & ""
                        .Session("PERMIS_HIST") = DR("PERMIS_HIST") & ""
                        .Session("EMAIL") = DR("USER_EMAIL") & ""

                        DR2 = GetDR(DAL.QueryData("SELECT * FROM SYS_LOGS WHERE CATEGORY='Log On' AND UPPER(USER_NAME)='" & DR("USER_NAME").ToString.ToUpper & "' ORDER BY TRANS_DATE DESC"))

                        If Not IsNothing(DR2) Then
                            .Session("LAST_ACTION") = FormatDate(DR2("TRANS_DATE"), "DD/MM/YYYY HH:MIN") & " IP Addr : " & DR2("IP_ADDRESS") & ""
                            ClearObject(DR2)
                        End If

                        .Session("UID") = DataEncrypt(RandomData() & Format(Now, "ddMMyy hh:mm:ss"))
                        For I = 1 To 30
                            .Session("TASK" & I) = CanDo(I, actView)
                            If I = 6 OrElse (I >= 12 AndAlso I <= 16) OrElse I = 18 OrElse I = 20 Then
                                .Session("TASK" & I & "ADD") = CanDo(I, actModify)
                            End If
                        Next
                        ''Aoy 11/05/2552
                        If DR("CHG_PWD_DATE") & "" = "" OrElse (DR("PWD_EXPIRE_DATE") & "" <> "" AndAlso (CInt("0" & DR("DAY_EXPIRE")) > 0 And CDate(DR("PWD_EXPIRE_DATE")) < Today)) Then
                            Msg = "PWD_EXPIRE"
                            Exit Try
                        End If
                    End With
                End If
            End If

            If Msg = "Invalid user name or password." Then
                BLL.InsertAudit("Log On Failed", "Try to Log On '" & UserName & "'", "")
            End If
        Catch ex As Exception
            Msg = GetErrorMsg(ex, "", "LOGON")
        Finally
            ClearObject(DT)
        End Try

        Return Msg
    End Function

    Public Sub ClearSessionData(Optional ByVal NotClearType As String = "")
        Try
            With HttpContext.Current
                If NotClearType <> "QUESTIONNAIRE" Then ClearObject(.Session("Q_DATA"))
                If NotClearType <> "MASTER" Then ClearObject(.Session("S_MDATA")) : ClearObject(.Session("S_DCCONFIG2")) : ClearObject(.Session("S_SEARCHCONFIG"))
                If NotClearType <> "SEARCH" Then ClearObject(.Session("S_DATA"))
                If NotClearType <> "SUPPORT" Then ClearObject(.Session("SUP_DATA"))
            End With
        Catch ex As Exception
        End Try
    End Sub

    'Aoy 27/03/2552
    'Created By Aoy 27/03/2552 Updated By Aoy 27/03/2552
    Public Function GenAddress(Optional ByVal HouseNo As String = "", Optional ByVal Village As String = "", Optional ByVal Moo As String = "", Optional ByVal Soi As String = "", Optional ByVal Road As String = "", Optional ByVal Tumbon As String = "", Optional ByVal Amphur As String = "", Optional ByVal Province As String = "", Optional ByVal ZipCode As String = "", Optional ByVal Country As String = "", Optional ByVal Lang As String = "ENG") As String
        Dim Address As String = ""
        If HouseNo <> "" Then
            Address &= HouseNo & " "
        End If
        If Village <> "" Then
            Address &= Village & " "
        End If
        If Moo <> "" Then
            If Lang = "ENG" Then
                Address &= "Moo " & Moo & " "
            Else
                Address &= "หมู่ที่ " & Moo & " "
            End If
        End If
        If Soi <> "" Then
            If Lang = "ENG" Then
                Address &= "Soi " & Soi & " "
            Else
                Address &= "ซอย " & Soi & " "
            End If
        End If
        If Road <> "" Then
            If Lang = "ENG" Then
                Address &= Road & " Road. "
            Else
                Address &= "ถนน " & Road & " "
            End If
        End If
        If Tumbon <> "" Then
            If Lang <> "ENG" Then
                Address &= "ต. "
            End If
            Address &= Tumbon & " "
        End If
        If Amphur <> "" Then
            If Lang <> "ENG" Then
                Address &= "อ. "
            End If
            Address &= Amphur & " "
        End If
        If Province <> "" Then
            If Lang <> "ENG" Then
                Address &= "จ. "
            End If
            Address &= Province & " "
        End If
        If ZipCode <> "" Then
            Address &= ZipCode & " "
        End If
        If Country <> "" Then
            If Lang <> "ENG" Then
                Address &= "ประเทศ"
            End If
            Address &= Country & " "
        End If
        Return Address
    End Function



    Public Sub LoadAgeCombo(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal BlankText As String = "")
        Dim I As Integer

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If
        For I = 15 To 35
            C.Items.Add(New ListItem(I, I))
        Next
        C.Items.Add(New ListItem("Over 35", "99"))

    End Sub

#Region "FCKEditor"
    Public Function ReplacePath(ByVal Action As String, ByVal str As String) As String
        Dim cURL As System.Uri = Nothing
        Dim strFind(0) As String
        Dim strReplace(0) As String
        Dim i As Integer

        Try
            Select Case Action
                Case "SAVE"
                    strFind(0) = "/" & ProjectName & "/"
                    strReplace(0) = "../"
                    strFind(1) = "../../"
                    strReplace(1) = "../"

                    For i = 0 To strFind.Length - 1
                        str = Replace(str, strFind(i), strReplace(i))
                    Next
                Case "LOAD"
                    strFind(0) = "../"
                    strReplace(0) = "/" & ProjectName & "/"

                    For i = 0 To strFind.Length - 1
                        str = Replace(str, strFind(i), strReplace(i))
                    Next
                Case "SEND_EMAIL"
                    cURL = HttpContext.Current.Request.Url

                    strFind(0) = "../"
                    'strReplace(0) = cURL.Scheme & cURL.SchemeDelimiter.ToString() & cURL.Host & "/" & ProjectName & "/"
                    strReplace(0) = "http://" & cURL.Host & "/" & ProjectName & "/"

                    For i = 0 To strFind.Length - 1
                        str = Replace(str, strFind(i), strReplace(i))
                    Next
            End Select
        Catch ex As Exception
        Finally
            ClearObject(strFind) : ClearObject(strReplace)
        End Try

        Return str
    End Function

    Public Sub DeleteFCKEditorUploadFile(ByVal Prefix As String, ByVal sPath As String)
        Dim tFolder As IO.DirectoryInfo = Nothing
        Dim tFiles As IO.FileInfo() = Nothing
        Dim obj As IO.FileInfo
        Dim tPath As String = "\" & ProjectName & "\" & "Files\"
        Dim PathFile, FileName As String

        Try
            If Prefix <> "" Then
                'Image
                PathFile = HttpContext.Current.Server.MapPath(tPath & sPath & ImgPath & "\")
                tFolder = New IO.DirectoryInfo(PathFile)
                If tFolder.Exists Then
                    tFiles = tFolder.GetFiles()
                    For Each obj In tFiles
                        FileName = obj.Name
                        If FileName Like Prefix & "_*" Then
                            Try
                                DeleteFile(PathFile & FileName)
                            Catch tex As Exception
                            End Try
                        End If
                    Next
                End If
                ClearObject(tFolder) : ClearObject(tFiles)

                'Attach
                PathFile = HttpContext.Current.Server.MapPath(tPath & sPath & LinkPath & "\")
                tFolder = New IO.DirectoryInfo(PathFile)
                If tFolder.Exists Then
                    tFiles = tFolder.GetFiles()
                    For Each obj In tFiles
                        FileName = obj.Name
                        If FileName Like Prefix & "_*" Then
                            Try
                                DeleteFile(PathFile & FileName)
                            Catch tex As Exception
                            End Try
                        End If
                    Next
                End If
                ClearObject(tFolder) : ClearObject(tFiles)
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region

    'Created By Aoy 08/04/2552 Find RateID of Propersol
    Public Function GetRateID(ByVal PositionName As String) As String
        Dim RateID As String = "1"
        Dim ITPos As String = "IT,System Administrator,Programmer,Developer,Web Developer,Engineer"
        If ITPos.IndexOf(PositionName) >= 0 Then
            RateID = "2"
        End If
        Return RateID
    End Function

    Public Function FormatReportDate(ByVal d As Integer, ByVal m As Integer, ByVal y As Integer) As String
        If y > 2500 Then y -= 543
        Return m & "/" & d & "/" & y
    End Function

    Public Function FormatReportDate(ByVal d As String) As String
        Dim TmpDate As Object = AppDateValue(AppFormatSQLDate(d))
        If Not IsNothing(TmpDate) Then
            Return FormatReportDate(TmpDate.Day, TmpDate.Month, TmpDate.Year)
        Else
            Return ""
        End If
    End Function

    Public Function GetSaveDataStructure(ByVal Type As String) As DataTable
        Dim DT As New DataTable

        'DT.Columns.Add("XXX", System.Type.GetType("System.String"))
        'DT.Columns.Add("XXX", System.Type.GetType("System.DateTime"))
        'DT.Columns.Add("XXX", System.Type.GetType("System.Int32"))
        'DT.Columns.Add("XXX", System.Type.GetType("System.Double"))

        Try
            Select Case Type
                Case "MAIL_ATTACHMENT"
                    DT.Columns.Add("FILE_NAME")
                    DT.Columns.Add("FILE_PATH")
                    DT.Columns.Add("CAN_DELETE_FILE")

            End Select

            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 04/05/2552 Generate prefix by TypeCode format of PTTICT 
    Public Function GeneratePrefix(ByVal ProjectType As String) As String
        Dim Prefix As String
        Prefix = DAL.LookupSQL("SELECT SERVICE_PREFIX_NAME FROM REF_PROJECT_TYPES WHERE PROJECT_TYPE=" & ProjectType) & FormatDate(Date.Now, "YYMM")
        Return Prefix
    End Function

    Public Function GetStringListInSQL(ByVal DataList As String) As String
        Dim OneData() As String
        Dim I As Integer
        Dim ResultList As String = ""
        If DataList <> "" Then
            OneData = Split(DataList, ",")
            For I = 0 To OneData.Length - 1
                If ResultList <> "" Then ResultList += ","
                If OneData(I) <> "" Then
                    ResultList += "'" & OneData(I) & "'"
                End If
            Next
        End If
        Return ResultList
    End Function

    Public Function GetCtrl(ByVal ColumnType As String, ByVal Row As String, ByVal Value As String) As String
        Dim Ctrl As String = ""
        Dim DR As DataRow
        Dim CtrlType As String = ""
        Dim Param1 As String = ""
        Dim Param2 As String = ""
        Dim Param3 As String = ""
        Try
            DR = GetDR(DAL.SearchConfigColumnTypes(ColumnType))
            If Not IsNothing(DR) Then
                CtrlType = DR("CTRL_TYPE") & ""
                Param1 = DR("PARAM1") & ""
                Param2 = DR("PARAM2") & ""
                Param3 = DR("PARAM3") & ""
                Select Case CtrlType.ToUpper
                    Case "LABEL"
                End Select
            End If
        Catch ex As Exception
            Ctrl = ""
        End Try
        Return Ctrl
    End Function

    Public Function GetServiceStatus(ByVal ServiceStatus As String) As String
        Dim ServiceStatusList As String = ""
        Select Case ServiceStatus
            Case "1", "9" : ServiceStatusList = "2,6,7,8" 'New, Re-open (New)
            Case "2"
                If HttpContext.Current.Session("USER_TYPE") & "" = "1" Then
                    ServiceStatusList = "2,3,4,5,7" 'Assigned
                Else
                    ServiceStatusList = "3,4,5,7" 'Assigned
                End If
            Case "3"
                If HttpContext.Current.Session("USER_TYPE") & "" = "1" Then
                    ServiceStatusList = "2,4,5,7" 'Inprogress
                Else
                    ServiceStatusList = "4,5,7" 'Inprogress
                End If
            Case "4"
                'ServiceStatusList = "2,5,6,7" 'Feedback
                If HttpContext.Current.Session("USER_TYPE") & "" = "1" Then
                    ServiceStatusList = "2,4,5,6,7" 'Feedback
                Else
                    ServiceStatusList = "4,5,6,7" 'Feedback
                End If
            Case "5" : ServiceStatusList = "6,7,8" 'Resolved
            Case "6"
                'ServiceStatusList = "2,7,8" 'Pending
                If HttpContext.Current.Session("USER_TYPE") & "" = "1" Then
                    ServiceStatusList = "2,7,8" 'Pending
                Else
                    ServiceStatusList = "7,8" 'Pending
                End If
            Case "7"
                'ServiceStatusList = "2,3,5" 'Reject
                ServiceStatusList = "2,3,4" 'Reject
            Case "8" : ServiceStatusList = "9" 'Closed
        End Select
        Return ServiceStatusList
    End Function

    Public Function GetCssClassActionStatus(ByVal ServiceStatus As String) As String
        Dim CssClass As String = ""
        Select Case ServiceStatus
            Case "1", "9" : CssClass = "ActionNew"
            Case "2" : CssClass = "ActionAssign"
            Case "3" : CssClass = "ActionInprogress"
            Case "4" : CssClass = "ActionFeedback"
            Case "5" : CssClass = "ActionResolved"
            Case "6" : CssClass = "ActionPending"
            Case "7" : CssClass = "ActionReject"
            Case "8" : CssClass = "ActionClosed"
        End Select
        Return CssClass
    End Function

    'return Days Hours Minutes
    Public Function GenTime(ByVal Minute As String, Optional ByVal IsSLA As Boolean = False, Optional ByVal Lang As String = "EN") As String
        Dim Time As String = ""
        Dim iMinute As Integer = 0
        Dim rDay, rHour, rMinute As Integer
        If Minute <> "" Then
            iMinute = ToInt(Minute)
            If iMinute < 0 Then iMinute = -(iMinute)
            rDay = iMinute \ 1440
            rHour = (iMinute \ 60) Mod 24
            rMinute = iMinute Mod 60
            Time = rDay & IIf(Lang = "EN", " Days ", " วัน ") & rHour & IIf(Lang = "EN", " Hours ", " ชั่วโมง ") & rMinute & IIf(Lang = "EN", " Minutes ", " นาที ") & ""
        End If
        Return Time
    End Function

    Public Function GetDBValue(ByVal ValueList As String, ByVal strSplit As String, ByVal DataType As Integer) As String
        Dim ValueListR As String = ""
        Dim arrVal As String()
        Dim Val As String = ""
        Dim ValChk As String = ""
        If ValueList <> "" Then
            arrVal = ValueList.Split(strSplit)
            For Each Val In arrVal
                ValChk = DB.SQLValue(Val, DataType).ToString().Trim()
                If ValChk <> "" AndAlso ValChk <> "NULL" Then
                    ValueListR &= ValChk & strSplit
                End If
            Next
            If ValueListR <> "" Then ValueListR = Left(ValueListR, ValueListR.Length - 1)
        End If
        Return ValueListR
    End Function

    Public Function RandomData() As String
        Dim Data As String = ""
        Dim CapAlphabet As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
        Dim Alphabet As String = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
        Dim Symbolic As String = "1,2,3,4,5,6,7,8,9,0"
        Dim CapAlphabetArr As String() = CapAlphabet.Split(",")
        Dim AlphabetArr As String() = Alphabet.Split(",")
        Dim SymbolicArr As String() = Symbolic.Split(",")
        Dim Random As New Random()
        Data = CapAlphabetArr(Random.Next(0, 20)) + CapAlphabetArr(Random.Next(0, 20)) + AlphabetArr(Random.Next(0, 20)) + AlphabetArr(Random.Next(0, 20)) + SymbolicArr(Random.Next(0, 10)) + AlphabetArr(Random.Next(0, 20)) + SymbolicArr(Random.Next(0, 10)) + SymbolicArr(Random.Next(0, 10))
        Return Data
    End Function

    Public Function GetParamValue(ByVal sKeys As String) As String
        Dim request As String = String.Empty
        Try
            If Not IsNothing(HttpContext.Current.Request.Params(sKeys)) AndAlso HttpContext.Current.Request.Params(sKeys) & "" <> String.Empty Then
                request = HttpContext.Current.Request.Params(sKeys).ToString()
            End If
        Catch ex As Exception
            request = ""
        End Try
        Return request
    End Function

    Public Function GetDashboardCriteria(ByVal Status As String, ByVal ViewGroup As String) As String
        Dim Criteria As String = ""
        Dim UserName As String = ""
        Dim UserNameChk As String
        Dim GroupID As String = ""
        Dim GroupIDChk As String
        Dim DT As DataTable = Nothing
        Try
            UserName = HttpContext.Current.Session("USER_NAME") & ""
            GroupID = HttpContext.Current.Session("GROUP_ID") & ""
            Select Case Status
                Case "0" 'Over SLA
                    Criteria &= " SV_SERVICE_STATUS NOT IN (9,8) AND (OVER_SLA < 0) AND (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL "

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If
                Case "1" 'New (Unassign,9:ReNew)
                    Criteria = " SERVICE_STATUS IN (1,9) AND SV_SERVICE_STATUS <> 8 AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL"
                    Criteria &= " "

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then
                            Criteria &= " AND (UPPER(USER_NAME)=" & UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                        End If
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then
                            Criteria &= " AND (USER_GROUP_ID=" & GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                        End If
                    End If
                Case "2" 'Assigned
                    Criteria = " SV_SERVICE_STATUS <> 8 AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL"

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If

                Case "3" 'Recently Modified (3:Inprogress)
                    Criteria = " SERVICE_STATUS IN (3) AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL" '-- remove display in status 8:Closed --> SERVICE_STATUS IN (3,6,8)

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If

                Case "4" 'Response (4:Feedback,7:Reject)
                    Criteria = " SERVICE_STATUS IN (4,7) AND SV_SERVICE_STATUS <> 8 AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL"

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If

                Case "5" 'Resolved
                    Criteria = " SV_SERVICE_STATUS <> 8 AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL"
                    Criteria &= ""

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If
                Case "6" 'Pending
                    Criteria = " SV_SERVICE_STATUS IN (6) AND (OVER_SLA >= 0)  AND NOT (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL "
                    Criteria &= ""

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If
                Case "10" 'ใกล้ตก SLA
                    Criteria &= " SV_SERVICE_STATUS NOT IN (8) AND (OVER_SLA >= 0)  AND (OVER_SLA < RESOLUTION_TIME2 / 2) AND ASSIGN_TO_ACTION_ID IS NULL "

                    If ViewGroup = "VIEW_MYSELF" Then
                        UserNameChk = DB.SQLValue(UserName.ToUpper(), DBUTIL.FieldTypes.ftText).ToString().Trim()
                        If UserNameChk <> "" AndAlso UserNameChk <> "NULL" Then Criteria &= " AND (UPPER(USER_NAME)=" & _
                        UserNameChk & " OR UPPER(ASSIGN_BY)=" & UserNameChk & ")"
                    Else
                        GroupIDChk = DB.SQLValue(GroupID, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                        If GroupIDChk <> "" AndAlso GroupIDChk <> "NULL" Then Criteria &= " AND (USER_GROUP_ID=" & _
                        GroupIDChk & " OR ASSIGN_BY_GROUP_ID=" & GroupIDChk & ")"
                    End If
            End Select
        Catch ex As Exception
            'Msg = GetErrorMsg(ex, "LOAD")
        Finally
            ClearObject(DT)
        End Try

        Return Criteria
    End Function
End Module



