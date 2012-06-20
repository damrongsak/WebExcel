#Region ".NET Framework Class Import"
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Web
Imports System.IO
#End Region


Public Class BLLComponent

    Private DAL As DALComponent
    Private DAL2 As DALComponent

    Public Sub New()
        DAL = New DALComponent
        DAL2 = New DALComponent("2")
    End Sub

    Protected Overrides Sub Finalize()
        ClearObject(DAL)
        MyBase.Finalize()
    End Sub

#Region "Admin"
    Public Function DeleteUserData(ByVal UserName As String, ByVal KeyID As String, ByVal FilePath As String, ByVal Img As String) As String
        Dim ret As String
        Dim tmpPath As String

        ret = DAL.MngUserData(opDELETE, UserName)
        If ret = "" Then
            tmpPath = HttpContext.Current.Server.MapPath(gFilePath & FilePath)
            If Img <> "" Then
                Try
                    DeleteFile(tmpPath & Img)
                Catch tex1 As Exception
                End Try
            End If
            'DeleteFCKEditorUploadFile(KeyID, FilePath)
        End If
        Return ret
    End Function
#End Region

    Public Sub WriteSytemLog(ByVal LogType As String, ByVal Msg As String, Optional ByVal RefID1 As String = "", Optional ByVal RefID2 As String = "")
        Try
            DAL.InsertAudit(LogType, Msg, HttpContext.Current.Session("USER_NAME") & "", HttpContext.Current.Request.ServerVariables("LOCAL_ADDR") & "", RefID1, RefID2)
        Catch ex As Exception
            ' Ignore write log error
        End Try
    End Sub

    Public Sub InsertAudit(ByVal Category As String, ByVal Action As String, ByVal User As String, Optional ByVal RefID1 As String = "" _
    , Optional ByVal RefID2 As String = "", Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing)
        DAL.InsertAudit(Category, Action, User, HttpContext.Current.Request.UserHostAddress & "", RefID1, RefID2, Conn, Trans)
    End Sub


    Public Function SearchAudit(ByVal FromDate As String, ByVal ToDate As String, ByVal Category As String, ByVal Action As String, ByVal User As String, Optional ByVal OtherCriteria As String = "", Optional ByVal OrderSQL As String = "") As DataTable
        Return DAL.SearchAudit(FromDate, ToDate, Category, Action, User, OtherCriteria, OrderSQL)
    End Function

    Public Sub UpdateSiteGroupView(ByVal SiteID As String, Optional ByVal AssignTo As String = "", Optional ByVal GroupID As String = "")
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim GroupView As String = ""
        Try
            DT = DAL.QueryData("SELECT DECODE(SU.GROUP_ID,NULL,SA.USER_GROUP_ID,SU.GROUP_ID) AS GROUP_ID " & _
            "FROM SERVICE_ACTIONS SA,SYS_USERS SU,SERVICES SV WHERE SA.SERVICE_ID=SV.SERVICE_ID" & _
            " AND SA.USER_NAME=SU.USER_NAME(+) AND SV.SITE_ID='" & SiteID & "' GROUP BY DECODE(SU.GROUP_ID,NULL,SA.USER_GROUP_ID,SU.GROUP_ID)")
            For Each DR In DT.Rows
                GroupView &= DR("GROUP_ID") & ","
                ClearObject(DR)
            Next
            If GroupView <> "" Then GroupView = "," & GroupView
            DAL.MngSiteData(opUPDATE, SiteID, Nothing, GroupView:=GroupView)

            If AssignTo <> "" Then
                DR = GetDR(DAL.SearchUserList(AssignTo))
                If Not IsNothing(DR) Then
                    If DR("PERMIS_INFOS") & "" = "0" Then
                        DAL.MngGroupData(opUPDATE, DR("GROUP_ID") & "", PermisInfos:="1")
                    End If
                End If
            Else
                DR = GetDR(DAL.SearchGroup(GroupID))
                If Not IsNothing(DR) Then
                    If DR("PERMIS_INFOS") & "" = "0" Then
                        DAL.MngGroupData(opUPDATE, DR("GROUP_ID") & "", PermisInfos:="1")
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DT)
        End Try
    End Sub

    Public Sub SendTemplateEMail(ByVal Subject As String, ByVal Sender As String _
        , ByVal Receiver As String, ByVal CCMail As String, ByVal Message As String _
        , Optional ByVal BCCMail As String = "", Optional ByVal IsAutoMail As Boolean = False _
        , Optional ByVal SendID As String = "", Optional ByVal Attachments As DataTable = Nothing)
        Dim DT As DataTable = Nothing
        Dim AttachFile() As String = Nothing
        Dim Status As String

        Try
            'Insert Image Tag to email in order to acknowledge that receiver opene email or not.
            AttachFile = GetMailAttachFile(Message, Attachments)

            Status = SendEMailData(Subject, Message, Sender, Receiver, CCMail, Filename:=AttachFile)

            'If Not IsNothing(Attachments) Then
            '    For Each DR In Attachments.Rows
            '        If DR("CAN_DELETE_FILE") & "" <> "N" AndAlso DR("CAN_DELETE_FILE") & "" <> "YNA" Then
            '            Try
            '                DeleteFile(HttpContext.Current.Server.MapPath(DR("FILE_PATH")))
            '            Catch
            '            End Try
            '        End If
            '    Next
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetMailAttachFile(ByRef Message As String, ByRef Attachments As DataTable) As String()
        'Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing

        Dim ArrFile() As String = Nothing
        Dim FileList, FileNameList, FileType As String
        Dim Detail As String

        Dim FileName As String
        Dim i As Integer

        Try
            Detail = Message

            ' File Name From System & Web
            'DT = HttpContext.Current.Session("ATTACHMENT")
            FileList = ""
            If Not IsNothing(Attachments) Then
                For Each DR In Attachments.Rows
                    FileList &= "," & DR("FILE_PATH")
                Next
            End If

            ' File Name From FCK.value
            FileNameList = ""
            Do While InStr(Detail.ToLower(), "<img") > 0
                i = InStr(Detail.ToLower(), "src=")
                Detail = Detail.Substring(i + 4)
                i = InStr(Detail, """")
                If i = 0 Then i = InStr(Detail, "'")
                FileName = Detail.Substring(0, i - 1)
                If FileName <> "" Then
                    FileType = FileName.Substring(FileName.LastIndexOf("."), 4)
                    If InStr("|" & imgFileType & "|", "|" & FileType & "|") > 0 Then
                        If InStr(FileName, "http://") = 0 Then
                            'FCK Image : Attach File Name
                            FileNameList &= "," & FileName
                            FileList &= "," & FileName.Replace("/" & ProjectName, "..")
                        ElseIf InStr(FileName, "/" & ProjectName) > 0 Then
                            FileNameList &= "," & FileName
                            FileName = FileName.Substring(FileName.IndexOf("/" & ProjectName))
                            FileList &= "," & FileName.Replace("/" & ProjectName, "..")
                        ElseIf InStr(FileName, "../") > 0 Then
                            FileNameList &= "," & FileName
                            FileList &= "," & FileName
                        End If
                    End If
                End If
                Detail = Detail.Substring(i)
            Loop

            ' Format Image Attachment in Message
            If FileNameList <> "" Then
                If FileNameList.StartsWith(",") Then FileNameList = FileNameList.Substring(1)
                ArrFile = FileNameList.Split(",")
                For i = 0 To UBound(ArrFile)
                    FileName = ArrFile(i)
                    Message = Message.Replace(FileName, "cid:" & GetFileName(FileName.Replace("/", "\")))
                Next i
            End If
            If FileList <> "" Then
                If FileList.StartsWith(",") Then FileList = FileList.Substring(1)
                Return FileList.Split(",")
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw ex
        Finally
            'ClearObject(DT)
        End Try
    End Function

    Public Sub TaskSendRegisUser(ByVal UserName As String, ByVal Password As String)
        Dim DTRecp As DataTable = Nothing 'Recipient
        Dim DTTmp As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim Template, Recipient, Message, Subject As String

        Try
            DTTmp = DAL.SearchTemplate("3")
            DR = GetDR(DTTmp) : ClearObject(DTTmp)
            If Not IsNothing(DR) Then
                Template = HttpContext.Current.Server.HtmlDecode(DR("MAIL_TEMP_DETAIL") & "")
                Template = ReplacePath("LOAD", Template)
                Subject = DR("MAIL_TEMP_SUBJECT") & ""

                DTRecp = DAL.SearchUserList(UserName)
                DR = GetDR(DTRecp)
                If Not IsNothing(DR) Then
                    Message = ""
                    Recipient = DR("USER_EMAIL") & ""
                    If Recipient <> "" Then
                        Message = Template
                        Message = Message.Replace("[DATE_UPDATED_T]", FormatDate(DR("DATE_UPDATED"), "ÇÑ¹·Õè DD à´×Í¹´´´´ ¾.È.»»»»"))
                        Message = Message.Replace("[USER_DESC]", DR("USER_DESC") & "")
                        Message = Message.Replace("[USER_NAME]", DR("USER_NAME") & "")
                        Message = Message.Replace("[PASSWORD]", Password)
                        Message = Message.Replace("[EMAIL]", DR("USER_EMAIL") & "")
                        Message = Message.Replace("[DATE_CREATED_T]", FormatDate(DR("DATE_CREATED"), "DD à´×Í¹´´´´ ¾.È.»»»»"))
                        Message = Message.Replace("[DATE_UPDATED_E]", FormatDate(DR("DATE_UPDATED"), "DD/MM/YYYY"))
                        'Send Mail
                        Try
                            SendTemplateEMail(Subject, "", Recipient, "", Message)
                        Catch
                        End Try

                    End If
                End If 'If Not IsNothing(DTRecp) AndAlso DTRecp.Rows.Count > 0 Then
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DTRecp) : ClearObject(DTTmp)
        End Try
    End Sub

    Public Sub TaskSendService1(ByVal ServiceID As String, ByVal Remail As String)
        Dim DTRecp As DataTable = Nothing 'Recipient
        Dim DTTmp As DataTable = Nothing
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim Template, Recipient, Message, Subject, ProbDetail As String

        ProbDetail = ""
        Try
            DTTmp = DAL.SearchTemplate("5")
            DR = GetDR(DTTmp) : ClearObject(DTTmp)
            If Not IsNothing(DR) Then
                Template = HttpContext.Current.Server.HtmlDecode(DR("MAIL_TEMP_DETAIL") & "")
                Template = ReplacePath("LOAD", Template)
                Subject = DR("MAIL_TEMP_SUBJECT") & ""

                DTRecp = DAL.SearchServiceAction(ServiceID)
                DR = GetDR(DTRecp)
                If Not IsNothing(DR) Then
                    Subject = Subject.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                    Subject = Subject.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")
                    Message = ""
                    Recipient = Remail
                    If Recipient <> "" Then
                        Message = Template

                        Message = Message.Replace("[PROJECT_TYPE_DESC]", DR("PROJECT_TYPE_DESC") & "")
                        Message = Message.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                        Message = Message.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")
                        Message = Message.Replace("[ASSIGN_DATE_T]", FormatDate(DR("SERVICE_DATE"), "DD à´×Í¹´´´´ ¾.È.»»»»  HH:MIN ¹."))
                        Message = Message.Replace("[NOTE]", DR("ASSIGN_NOTE") & "")
                        Message = Message.Replace("[SITE_ID]", DR("SITE_ID") & "")
                        Message = Message.Replace("[SITE_NAME]", DR("SITE_NAME") & "")
                        Message = Message.Replace("[SITE_ADDRESS]", DR("ADDRESS") & "")

                        Message = Message.Replace("[INFORMER_NAME]", DR("INFORMER_NAME") & "")
                        Message = Message.Replace("[INFORMER_TEL]", DR("INFORMER_TEL") & "")
                        Message = Message.Replace("[CALL_DETAIL]", DR("CALL_DETAIL") & "")
                        Message = Message.Replace("[SEVERITY_LEVEL_DESC]", DR("SEVERITY_LEVEL_DESC") & "")
                        Message = Message.Replace("[RESOLUTION_TIME_T]", GenTime(DR("RESOLUTION_TIME") & "", Lang:="TH"))

                        'Send Mail
                        Try
                            SendTemplateEMail(Subject, "", Recipient, "", Message)
                        Catch
                        End Try

                    End If
                    ClearObject(DR)
                End If 'If Not IsNothing(DTRecp) AndAlso DTRecp.Rows.Count > 0 Then
                ClearObject(DR)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DTRecp) : ClearObject(DTTmp)
        End Try
    End Sub

    Public Sub TaskSendService(ByVal ActionID As String, ByVal ServiceID As String)
        Dim DTRecp As DataTable = Nothing 'Recipient
        Dim DTTmp As DataTable = Nothing
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim DR2 As DataRow = Nothing
        Dim DR3 As DataRow = Nothing
        Dim Template, Recipient, Message, Subject As String

        Try
            DTRecp = DAL.SearchServiceAction(ServiceID, ActionID)
            DR = GetDR(DTRecp)
            If Not IsNothing(DR) Then
                If DR("SERVICE_STATUS") & "" = "2" Then
                    DTTmp = DAL.SearchTemplate("4")
                Else
                    DTTmp = DAL.SearchTemplate("5")
                End If
                DR2 = GetDR(DTTmp) : ClearObject(DTTmp)
                If Not IsNothing(DR2) Then
                    If DR("USER_NAME") & "" <> "" Then
                        Recipient = DR("USER_EMAIL") & ""
                        Template = HttpContext.Current.Server.HtmlDecode(DR2("MAIL_TEMP_DETAIL") & "")
                        Template = ReplacePath("LOAD", Template)
                        Subject = DR2("MAIL_TEMP_SUBJECT") & ""
                        Subject = Subject.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                        Subject = Subject.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")
                        Message = ""
                        If Recipient <> "" Then
                            Message = Template
                            Message = Message.Replace("[PROJECT_TYPE_DESC]", DR("PROJECT_TYPE_DESC") & "")
                            Message = Message.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                            Message = Message.Replace("[NOTE]", DR("NOTE") & "")
                            Message = Message.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")

                            Message = Message.Replace("[SITE_ID]", DR("SITE_ID") & "")
                            Message = Message.Replace("[SITE_NAME]", DR("SITE_NAME") & "")
                            Message = Message.Replace("[SITE_ADDRESS]", DR("ADDRESS") & " " & DR("PROVINCE_NAME") & " " & DR("ZIP_CODE") & "")

                            Message = Message.Replace("[INFORMER_NAME]", DR("INFORMER_NAME") & "")
                            Message = Message.Replace("[INFORMER_TEL]", DR("INFORMER_TEL") & "")

                            Message = Message.Replace("[CALL_DETAIL]", DR("CALL_DETAIL") & "")
                            Message = Message.Replace("[SEVERITY_LEVEL_DESC]", DR("SEVERITY_LEVEL_DESC") & "")

                            If DR("USER_TYPE") & "" = "2" Then
                                Message = Message.Replace("[RESOLUTION_TIME_T]", GenTime(DR("RESOLUTION_TIME") & "", Lang:="TH"))
                                Message = Message.Replace("[RESOLUTION_TIME_E]", GenTime(DR("RESOLUTION_TIME") & "", Lang:="EN"))
                            Else
                                Message = Message.Replace("[RESOLUTION_TIME_T]", GenTime(DR("RESOLUTION_TIME2") & "", Lang:="TH"))
                                Message = Message.Replace("[RESOLUTION_TIME_E]", GenTime(DR("RESOLUTION_TIME2") & "", Lang:="EN"))
                            End If
                            Message = Message.Replace("[ASSIGN_DATE_T]", FormatDate(DR("SERVICE_DATE"), "DD à´×Í¹´´´´ ¾.È.»»»»  HH:MIN ¹."))
                            Message = Message.Replace("[ASSIGN_DATE_E]", FormatDate(DR("SERVICE_DATE"), "DD/MM/YYYY HH:MIN"))

                            If DR("SERVICE_STATUS") & "" = "2" OrElse DR("SERVICE_STATUS") & "" = "4" OrElse DR("SERVICE_STATUS") & "" = "7" Then
                                Message = Message.Replace("[PATH_REPLY]", "http://" & ProjectServer & "/" & ProjectName & _
                                "/Service/UpdateActionService.aspx?ActionID=" & DataEncrypt(ActionID) & "&ServiceID=" & _
                                DataEncrypt(ServiceID) & "&U=" & DataEncrypt(DR("USER_NAME") & ""))
                            ElseIf DR("SERVICE_STATUS") & "" = "5" AndAlso DR("RESPONSE_DATE") & "" = "" Then
                                Message = Message.Replace("[PATH_REPLY]", "http://" & ProjectServer & "/" & ProjectName & _
                                "/Service/UpdateActionService.aspx?ActionID=" & DataEncrypt(ActionID) & "&ServiceID=" & _
                                DataEncrypt(ServiceID) & "&U=" & DataEncrypt(DR("USER_NAME") & ""))
                            Else
                                Message = Message.Replace("[PATH_REPLY]", "")
                            End If

                            Message = Message.Replace("[REF_CALL_NUMBER]", String.Format("{0}", DR("REF_CALL_NUMBER")))

                            'Send Mail
                            Try
                                SendTemplateEMail(Subject, "", Recipient, "", Message)
                            Catch
                            End Try
                        Else
                            DTRecp = DAL.SearchUserList("", UserGroup:=DR("USER_GROUP_ID") & "")
                            For Each DR3 In DTRecp.Rows
                                Recipient = DR3("USER_EMAIL") & ""
                                Template = HttpContext.Current.Server.HtmlDecode(DR2("MAIL_TEMP_DETAIL") & "")
                                Template = ReplacePath("LOAD", Template)
                                Subject = DR2("MAIL_TEMP_SUBJECT") & ""
                                Subject = Subject.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                                Subject = Subject.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")
                                Message = ""
                                If Recipient <> "" Then
                                    Message = Template
                                    Message = Message.Replace("[PROJECT_TYPE_DESC]", DR("PROJECT_TYPE_DESC") & "")
                                    Message = Message.Replace("[SERVICE_NO]", DR("SERVICE_NO") & "")
                                    Message = Message.Replace("[NOTE]", DR("NOTE") & "")
                                    Message = Message.Replace("[SERVICE_STATUS_DESC]", DR("SERVICE_STATUS_DESC") & "")

                                    Message = Message.Replace("[SITE_ID]", DR("SITE_ID") & "")
                                    Message = Message.Replace("[SITE_NAME]", DR("SITE_NAME") & "")
                                    Message = Message.Replace("[SITE_ADDRESS]", DR("ADDRESS") & " " & DR("PROVINCE_NAME") & " " & DR("ZIP_CODE") & "")

                                    Message = Message.Replace("[INFORMER_NAME]", DR("INFORMER_NAME") & "")
                                    Message = Message.Replace("[INFORMER_TEL]", DR("INFORMER_TEL") & "")

                                    Message = Message.Replace("[CALL_DETAIL]", DR("CALL_DETAIL") & "")
                                    Message = Message.Replace("[SEVERITY_LEVEL_DESC]", DR("SEVERITY_LEVEL_DESC") & "")

                                    If DR("USER_TYPE") & "" = "2" Then
                                        Message = Message.Replace("[RESOLUTION_TIME_T]", GenTime(DR("RESOLUTION_TIME") & "", Lang:="TH"))
                                        Message = Message.Replace("[RESOLUTION_TIME_E]", GenTime(DR("RESOLUTION_TIME") & "", Lang:="EN"))
                                    Else
                                        Message = Message.Replace("[RESOLUTION_TIME_T]", GenTime(DR("RESOLUTION_TIME2") & "", Lang:="TH"))
                                        Message = Message.Replace("[RESOLUTION_TIME_E]", GenTime(DR("RESOLUTION_TIME2") & "", Lang:="EN"))
                                    End If
                                    Message = Message.Replace("[ASSIGN_DATE_T]", FormatDate(DR("SERVICE_DATE"), "DD à´×Í¹´´´´ ¾.È.»»»»  HH:MIN ¹."))
                                    Message = Message.Replace("[ASSIGN_DATE_E]", FormatDate(DR("SERVICE_DATE"), "DD/MM/YYYY HH:MIN"))

                                    If DR("SERVICE_STATUS") & "" = "2" OrElse DR("SERVICE_STATUS") & "" = "4" OrElse DR("SERVICE_STATUS") & "" = "7" Then
                                        Message = Message.Replace("[PATH_REPLY]", "http://" & ProjectServer & "/" & ProjectName & _
                                        "/Service/UpdateActionService.aspx?ActionID=" & DataEncrypt(ActionID) & "&ServiceID=" & _
                                        DataEncrypt(ServiceID) & "&U=" & DataEncrypt(DR("USER_NAME") & "") & "&G=" & DataEncrypt(DR("USER_GROUP_ID") & ""))
                                    ElseIf DR("SERVICE_STATUS") & "" = "5" AndAlso DR("RESPONSE_DATE") & "" = "" Then
                                        Message = Message.Replace("[PATH_REPLY]", "http://" & ProjectServer & "/" & ProjectName & _
                                        "/Service/UpdateActionService.aspx?ActionID=" & DataEncrypt(ActionID) & "&ServiceID=" & _
                                        DataEncrypt(ServiceID) & "&U=" & DataEncrypt(DR("USER_NAME") & "") & "&G=" & DataEncrypt(DR("USER_GROUP_ID") & ""))
                                    Else
                                        Message = Message.Replace("[PATH_REPLY]", "")
                                    End If

                                    Message = Message.Replace("[REF_CALL_NUMBER]", String.Format("{0}", DR("REF_CALL_NUMBER")))

                                    'Send Mail
                                    Try
                                        SendTemplateEMail(Subject, "", Recipient, "", Message)
                                    Catch
                                    End Try
                                End If
                            Next
                        End If
                    End If
                    ClearObject(DR2)
                End If
                ClearObject(DR)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DTRecp) : ClearObject(DTTmp)
        End Try
    End Sub

    Public Sub TaskSendRequestNewPass(ByVal UserName As String)
        Dim DTRecp As DataTable = Nothing 'Recipient
        Dim DTTmp As DataTable = Nothing
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim DR2 As DataRow = Nothing
        Dim Template, Recipient, Message, Subject As String
        Try
            DTRecp = DAL.SearchUserList(UserName)
            DR = GetDR(DTRecp)
            If Not IsNothing(DR) Then
                DTTmp = DAL.SearchTemplate("6")
                DR2 = GetDR(DTTmp) : ClearObject(DTTmp)
                If Not IsNothing(DR2) Then
                    Template = HttpContext.Current.Server.HtmlDecode(DR2("MAIL_TEMP_DETAIL") & "")
                    Template = ReplacePath("LOAD", Template)
                    Subject = DR2("MAIL_TEMP_SUBJECT") & ""
                    Message = ""
                    Recipient = gSender_EMail
                    If Recipient <> "" Then
                        Message = Template
                        Message = Message.Replace("[USER_DESC]", DR("USER_DESC") & "")
                        Message = Message.Replace("[USER_NAME]", DR("USER_NAME") & "")
                        Message = Message.Replace("[REQUEST_DATETIME]", FormatDate(Now, "DD/MM/YYYY HH:MIN"))
                        'Send Mail
                        Try
                            SendTemplateEMail(Subject, "", Recipient, "", Message)
                        Catch
                        End Try

                    End If
                    ClearObject(DR2)
                End If
                ClearObject(DR)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DTRecp) : ClearObject(DTTmp)
        End Try
    End Sub

    Public Sub TaskImportAMZSite()
        Dim dbConn As OleDb.OleDbConnection = Nothing
        Dim dbTrans As OleDb.OleDbTransaction = Nothing
        Dim dbConn2 As OleDb.OleDbConnection = Nothing
        Dim dbTrans2 As OleDb.OleDbTransaction = Nothing
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim DR2 As DataRow = Nothing
        Dim SQL As String = ""
        Dim op As Integer = opINSERT
        Dim ProvinceID As String = ""
        Try
            dbConn = DAL.OpenConn()
            dbTrans = DAL.BeginTrans(dbConn)
            dbConn2 = DAL2.OpenConn()
            dbTrans2 = DAL2.BeginTrans(dbConn2)
            SQL = "select 'AMZ' + lv.ProductLevelCode as SITE_ID, lv.ProductLevelName as SITE_NAME, '' as SITE_NAME2, '' as SITE_DESC, '' as SITE_TYPE, pf.CompanyAddress1 + ' ' + pf.CompanyAddress2 + ' ' + pf.CompanyCity as ADDRESS, pf.CompanyProvince as AmzProvinceID, '' as SALE_AREA, pf.CompanyTaxID as TAX_ID, pf.CompanyName as OWNER_NAME, '' as CONTACT_NAME, pf.CompanyTelephone as TEL_NO, pf.CompanyFax as FAX_NO, '' as EMAIL, '' as POS_SYSTEM_ID, '' as BACK_OFFICE_ID, '' as PLAN_INSTALL_DATE, '' as PLAN_OPEN_DATE, '' as INSTALL_DATE, '' as OPEN_DATE, '' as SLA_PROFILE_ID, '' as SAP_PLANT_CODE, '' as SAP_NAME, '' as SAP_SOLD_TO, '' as SAP_SHIP_TO, '' as DATE_UPDATED, 'ADMIN' as USER_UPDATED, '' as LATITUDE, '' as LONGITUDE, '' as BRAND_CODE, '3' as PROJECT_TYPE, '' as REMARK, '' as BRANCH_ID, pf.CompanyZipCode as ZIP_CODE, '' as CLOSED_DATE, '1' as STATUS, '' as COST_CENTER, '' as SAP_SITE_ID1, '' as SAP_SITE_ID2, '' as GROUP_VIEWS from computeraccess map, companyprofile pf, ProductLevel lv where (map.ProductLevelID = pf.CompanyID) and map.ProductLevelID = lv.ProductLevelID"
            DT = DAL.QueryData(SQL, dbConn2, dbTrans2)

            For Each DR In DT.Rows
                DR2 = GetDR(DAL.SearchSiteData(DR("SITE_ID") & "", Conn:=dbConn, Trans:=dbTrans))
                If Not IsNothing(DR2) Then op = opUPDATE
                If DR("AmzProvinceID") & "" <> "" Then
                    ProvinceID = DAL.LookupSQL("SELECT PROVINCE_ID FROM SAP_PROVINCE_MAPPING WHERE AMZ_PROVINCE_ID=" & DR("AmzProvinceID") & "", Conn:=dbConn, Trans:=dbTrans)
                End If

                DAL.MngSiteData(op, DR("SITE_ID") & "", DR("SITE_ID") & "", DR("SITE_NAME") & "", DR("SITE_NAME2") & "", DR("SITE_DESC") & "", DR("SITE_TYPE") & "", DR("ADDRESS") & "", ProvinceID, DR("SALE_AREA") & "", DR("TAX_ID") & "", DR("OWNER_NAME") & "", DR("CONTACT_NAME") & "", DR("TEL_NO") & "", DR("FAX_NO") & "", DR("EMAIL") & "", DR("POS_SYSTEM_ID") & "", DR("BACK_OFFICE_ID") & "", Status:=DR("STATUS") & "", PlanInstallDate:=FormatDate(DR("PLAN_INSTALL_DATE"), "DD/MM/YYYY"), PlanOpenDate:=FormatDate(DR("PLAN_OPEN_DATE"), "DD/MM/YYYY"), InstallDate:=FormatDate(DR("INSTALL_DATE"), "DD/MM/YYYY"), OpenDate:=FormatDate(DR("OPEN_DATE"), "DD/MM/YYYY"), SLAProfileID:=DR("SLA_PROFILE_ID") & "", SAPPlantCode:=DR("SAP_PLANT_CODE") & "", SAPName:=DR("SAP_NAME") & "", SAPSoldTo:=DR("SAP_SOLD_TO") & "", SAPShipTo:=DR("SAP_SHIP_TO") & "", Latitude:=DR("LATITUDE") & "", Longitude:=DR("LONGITUDE") & "", BranchCode:=DR("BRANCH_ID") & "", ProjectType:=DR("PROJECT_TYPE") & "", Remark:=DR("REMARK") & "", CostCenter:=DR("COST_CENTER") & "", CloseDate:=FormatDate(DR("CLOSED_DATE"), "DD/MM/YYYY"), SAPSiteID1:=DR("SAP_SITE_ID1") & "", SAPSiteID2:=DR("SAP_SITE_ID2") & "", BranchID:=DR("BRANCH_ID") & "", Conn:=dbConn, Trans:=dbTrans)
            Next
            DAL.CommitTrans(dbTrans)
            DAL2.CommitTrans(dbTrans2)
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DT)
            If Not IsNothing(dbTrans) Then DAL.RollbackTrans(dbTrans)
            If Not IsNothing(dbTrans2) Then DAL2.RollbackTrans(dbTrans2)
            DAL.CloseConn(dbConn)
            DAL2.CloseConn(dbConn2)
        End Try
    End Sub

#Region "EXPORT EXCEL"
    Public Function GenGridviewToTable(ByVal gv As GridView) As Table
        Try
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.ContentType = "application/octet-stream"
            Dim table As New Table()

            ' add the header row to the table 

            If gv.HeaderRow IsNot Nothing Then


                BLL.ExportControl(gv.HeaderRow)


                table.Rows.Add(gv.HeaderRow)
            End If

            ' add each of the data rows to the table 

            For Each row As GridViewRow In gv.Rows


                BLL.ExportControl(row)


                table.Rows.Add(row)
            Next

            ' add the footer row to the table 

            If gv.FooterRow IsNot Nothing Then


                BLL.ExportControl(gv.FooterRow)


                table.Rows.Add(gv.FooterRow)
            End If
            Return table

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ExportExcelByTable(ByVal fileName As String, ByVal TB As Table, Optional ByVal TB2 As Table = Nothing)
        Try

            HttpContext.Current.Response.Clear()

            'Export will take two parameter first one the name of Excel File, and second one for gridview to be exported 

            HttpContext.Current.Response.AddHeader("content-disposition", String.Format("attachment; filename={0}", fileName))

            HttpContext.Current.Response.ContentType = "application/octet-stream"

            Using strWriter As New StringWriter()

                Using htmlWriter As New HtmlTextWriter(strWriter)
                    ' render Gridlines

                    'TB.GridLines = GridLines.Both

                    ' render the table into the htmlwriter
                    htmlWriter.WriteLine("<style>" & _
                                         ".number {mso-number-format:0\.00; }" & _
                                         ".text {mso-number-format:General; text-align:general; }" & _
                                         ".num2text {mso-style-parent:text; mso-number-format:'\@';white-space:normal}" & _
                                         "</style>")


                    TB.RenderControl(htmlWriter)
                    If Not IsNothing(TB2) Then
                        TB2.RenderControl(htmlWriter)
                    End If
                    ' render the htmlwriter into the response 

                    HttpContext.Current.Response.Write(strWriter.ToString())


                    HttpContext.Current.Response.[End]()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return 0
    End Function


    Public Function ExportExcel(ByVal fileName As String, ByVal gv As GridView)
        Try
            HttpContext.Current.Response.Clear()

            'Export will take two parameter first one the name of Excel File, and second one for gridview to be exported 

            HttpContext.Current.Response.AddHeader("content-disposition", String.Format("attachment; filename={0}", fileName))

            HttpContext.Current.Response.ContentType = "application/octet-stream"

            Using strWriter As New StringWriter()


                Using htmlWriter As New HtmlTextWriter(strWriter)


                    ' Create a form to contain the grid 

                    Dim table As New Table()

                    ' add the header row to the table 

                    If gv.HeaderRow IsNot Nothing Then


                        BLL.ExportControl(gv.HeaderRow)


                        table.Rows.Add(gv.HeaderRow)
                    End If

                    ' add each of the data rows to the table 

                    For Each row As GridViewRow In gv.Rows


                        BLL.ExportControl(row)


                        table.Rows.Add(row)
                    Next

                    ' add the footer row to the table 

                    If gv.FooterRow IsNot Nothing Then


                        BLL.ExportControl(gv.FooterRow)


                        table.Rows.Add(gv.FooterRow)
                    End If


                    ' render Gridlines

                    table.GridLines = GridLines.Both

                    ' render the table into the htmlwriter

                    table.RenderControl(htmlWriter)

                    ' render the htmlwriter into the response 

                    HttpContext.Current.Response.Write(strWriter.ToString())


                    HttpContext.Current.Response.[End]()

                End Using

            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return 0
    End Function

    ''' Replace controls with literals 

    Public Function ExportControl(ByVal control As Control)


        For i As Integer = 0 To control.Controls.Count - 1


            Dim current As Control = control.Controls(i)

            If TypeOf current Is LinkButton Then


                control.Controls.Remove(current)


                control.Controls.AddAt(i, New LiteralControl(TryCast(current, LinkButton).Text))

            ElseIf TypeOf current Is ImageButton Then


                control.Controls.Remove(current)


                control.Controls.AddAt(i, New LiteralControl(TryCast(current, ImageButton).AlternateText))

            ElseIf TypeOf current Is HyperLink Then


                control.Controls.Remove(current)


                control.Controls.AddAt(i, New LiteralControl(TryCast(current, HyperLink).Text))

            ElseIf TypeOf current Is DropDownList Then


                control.Controls.Remove(current)


                control.Controls.AddAt(i, New LiteralControl(TryCast(current, DropDownList).SelectedItem.Text))

            ElseIf TypeOf current Is CheckBox Then


                control.Controls.Remove(current)


                control.Controls.AddAt(i, New LiteralControl(If(TryCast(current, CheckBox).Checked, "True", "False")))
            End If

            'Like that you may convert any control to literals 

            If current.HasControls() Then



                BLL.ExportControl(current)

            End If

        Next
        Return 0
    End Function

#End Region
End Class


