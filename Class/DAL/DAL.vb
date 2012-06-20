#Region ".NET Framework Class Import"
Imports System.Data
Imports System.Web
#End Region

Public Class DALComponent

#Region "Internal member variables"
    Private _UseStoredProc As Boolean = True

    Private DB As New DBUTIL

    Private _DB_Provider As String
    Private _DB_UserName As String
    Private _DB_Password As String
    Private _DB_DataSource As String
    Private _DB_Name As String
    Private _ConnectStr As String

    Private _DB_Provider2 As String
    Private _DB_UserName2 As String
    Private _DB_Password2 As String
    Private _DB_DataSource2 As String
    Private _DB_Name2 As String
    Private _ConnectStr2 As String

    Public ConnectStr As String
    Public ConnectStr2 As String
#End Region

#Region "Initialization"
    Public Sub New(Optional ByVal dbid As String = "")
        ReadDALConfigurations(dbid)
        DB.ConnectStr = _ConnectStr
        DB.DB_Provider = _DB_Provider
    End Sub

    Protected Overrides Sub Finalize()
        ClearObject(DB)
        MyBase.Finalize()
    End Sub

    Public Sub ReadDALConfigurations(Optional ByVal dbid As String = "")
        Dim Encrypt As New SecurityUtil

        Try
            _DB_Provider = ConfigurationManager.AppSettings("DB_Provider" & dbid)
            _DB_DataSource = ConfigurationManager.AppSettings("DB_DataSource" & dbid)
            _DB_Name = ConfigurationManager.AppSettings("DB_Name" & dbid)
            _DB_UserName = Encrypt.DecryptData(ConfigurationManager.AppSettings("DB_UserName" & dbid) & "") & ""
            _DB_Password = Encrypt.DecryptData(ConfigurationManager.AppSettings("DB_Password" & dbid) & "") & ""

            _ConnectStr = "Provider=" & _DB_Provider & ";Data Source=" & _DB_DataSource & ";User ID=" & _DB_UserName & ";Password=" & _DB_Password
            If _DB_Name <> "" Then _ConnectStr += ";Initial Catalog=" & _DB_Name

        Catch ex As Exception
            ClearObject(Encrypt)
        End Try
    End Sub
#End Region

#Region "General"

    Public Function OpenConn(Optional ByVal ConnectStr As String = "") As OleDb.OleDbConnection
        If ConnectStr = "" Then ConnectStr = _ConnectStr
        Return DB.OpenConn(ConnectStr)
    End Function

    Public Sub CloseConn(ByRef Conn As OleDb.OleDbConnection)
        DB.CloseConn(Conn)
    End Sub

    Public Function BeginTrans(ByRef Conn As OleDb.OleDbConnection) As OleDb.OleDbTransaction
        Return DB.BeginTrans(Conn)
    End Function

    Public Sub CommitTrans(ByRef Trans As OleDb.OleDbTransaction)
        DB.CommitTrans(Trans)
    End Sub

    Public Sub RollbackTrans(ByRef Trans As OleDb.OleDbTransaction)
        DB.RollbackTrans(Trans)
    End Sub

    Public Function QueryData(ByVal SQL As String, _
                             Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                             Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable

        Dim DT As New DataTable

        Try
            If Not IsNothing(Conn) Then
                DB.OpenDT(DT, SQL, Conn, Trans)
            Else
                DB.OpenDT(DT, SQL)
            End If

            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    Public Function ExecSQL(ByVal SQL As String, _
                             Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                             Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim DT As New DataTable

        Try
            If Not IsNothing(Conn) Then
                DB.ExecSQL(SQL, Conn, Trans)
            Else
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    Public Function LookupSQL(ByVal SQL As String, _
                    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Object
        Return DB.LookupSQL(SQL, Conn, Trans)
    End Function


    Public Function GenerateID(ByVal TableName As String, ByVal IDField As String, _
                     Optional ByVal Prefix As String = "", Optional ByVal IDLength As Integer = 0, _
                     Optional ByVal usrCriteria As String = "", _
                     Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                     Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Object

        Dim SQL As String
        Dim Criteria As String = ""
        Dim ID, tmp As Object

        Try
            SQL = "SELECT MAX(" & IDField & ") FROM " & TableName
            If usrCriteria <> "" Then Criteria += " AND " + usrCriteria
            If Prefix <> "" Then Criteria += " AND " + IDField + " LIKE '" & Prefix & "%'"
            If Criteria <> "" Then SQL += " WHERE " + Mid(Criteria, 6)
            ID = DB.LookupSQL(SQL, Conn, Trans)

            If ID & "" <> "" Then
                If InStr(ID, "-") > 0 Then
                    tmp = Mid(ID, 5, 4)
                    tmp = tmp + 1
                    ID = CStr(tmp).PadLeft(4, "0")
                Else
                    If Prefix <> "" Then ID = Mid(ID, Prefix.Length + 1)
                    ID = CLng(ID) + 1
                End If
            Else
                ID = 1
            End If
            If Prefix <> "" OrElse IDLength > 0 Then
                ID = Prefix & CStr(ID).PadLeft(IDLength, "0")
            End If

            Return ID
        Catch ex As Exception
            Throw New DALException(ex)
        End Try
    End Function

    Public Function GetMaxValue(ByVal usrTable As String, ByVal FieldName As String, ByVal usrSQL As String) As Object
        Try
            GetMaxValue = DB.GetMaxData(usrTable, FieldName, usrSQL)
        Catch ex As Exception
            Throw New DALException(ex)
        End Try
    End Function

    '21-05-52 ไว้ออกรายงานในกรณีเลือก Weekly
    Public Function GetMinMaxValue(Optional ByVal YY As String = "", Optional ByVal MM As String = "", Optional ByVal Week As String = "", Optional ByVal OtherCriteria As String = "", Optional ByVal GroupBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "REP_YEAR", YY, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "REP_MONTH", MM, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "REP_WEEK", Week, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT MIN(REP_DATE) AS MIN_DATE, MAX(REP_DATE) AS MAX_DATE,MAX(REP_YEAR) AS REP_YEAR, MAX(REP_MONTH) AS REP_MONTH, MAX(REP_WEEK) AS REP_WEEK FROM CALENDARS "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria

            If GroupBy <> "" Then
                SQL &= " GROUP BY " & GroupBy
            Else
                SQL &= " GROUP BY REP_DATE"
            End If

            SQL &= " ORDER BY REP_WEEK"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SQLDate(ByVal D As Object) As String
        Return DB.SQLDate(D)
    End Function

    Public Function SQLDateTime(ByVal DT As Object) As String
        Return DB.SQLDateTime(DT)
    End Function

    Public Function GetSQLValue(ByVal SQL As String) As Object
        Dim Conn As OleDb.OleDbConnection = Nothing

        Try
            Conn = DB.OpenConn(_ConnectStr)
            GetSQLValue = DB.LookupSQL(SQL)
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            DB.CloseConn(Conn)
        End Try
    End Function

    Public Sub AddCaseCriteria(ByRef CaseCriteria As String, ByVal ChkCriteria As String, ByVal FieldName As String, Optional ByVal TrueValue As String = "", Optional ByVal FalseValue As String = "")
        Dim CaseSQL As String = ""
        If ChkCriteria <> "" Then
            CaseSQL = " CASE WHEN " & ChkCriteria & " THEN " & TrueValue & " ELSE " & FalseValue & " END AS " & FieldName
        End If
        If ChkCriteria <> "" Then
            If CaseCriteria <> "" Then CaseCriteria += ","
            CaseCriteria += CaseSQL
        End If
    End Sub

    Public Sub SetCriteriaArray(ByVal CriteriaList As String, ByVal AllScore As String, ByRef CriteriaArray() As String, ByRef EachScore As Double)
        If CriteriaList <> "" Then
            CriteriaArray = Split(CriteriaList, ",")
            EachScore = ToInt(AllScore) / (CriteriaArray.Length)
        End If
    End Sub

    Public Function GetCodeList(ByVal RefTable As String, ByVal CodeField As String, ByVal DescField As String, ByVal DescValue As String, Optional ByVal CodeType As DBUTIL.FieldTypes = DBUTIL.FieldTypes.ftText, Optional ByVal IsCriteria As Boolean = False) As String
        Dim DT As DataTable
        Dim DR As DataRow
        Dim ResultList As String = ""
        Dim SQL As String = ""
        Try
            If IsCriteria = False Then
                SQL = "SELECT " & CodeField & " FROM " & RefTable & " WHERE " & DescField & " LIKE '%" & DescValue & "%'  ORDER BY " & CodeField
            Else
                SQL = "SELECT " & CodeField & " FROM " & RefTable & " WHERE " & DescValue & "   ORDER BY " & CodeField
            End If

            DT = QueryData(SQL)
            If Not IsNothing(DT) Then
                For Each DR In DT.Rows
                    If Not IsNothing(DR) Then
                        If DR(CodeField) & "" <> "" Then
                            If ResultList <> "" Then ResultList += ","
                            If CodeType = DBUTIL.FieldTypes.ftNumeric Then
                                ResultList += DR(CodeField) & ""
                            Else
                                ResultList += "'" & DR(CodeField) & "'"
                            End If
                        End If
                    End If
                Next
            End If
            Return ResultList
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Administrator"
    Public Function ChangePassword(ByVal UID As String, ByVal UPwd As String, ByVal UUpdated As String) As Object
        Dim SQL1, SQL2, SQL As String
        Dim op As Integer
        Dim result As Integer

        Try
            UID = UID.ToUpper()
            SQL = "" : SQL1 = "" : SQL2 = ""

            op = DBUTIL.opUPDATE

            If UPwd <> "****" Then DB.AddSQL(op, SQL1, SQL2, "PASSWORD", UPwd, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op, SQL1, SQL2, "CHG_PWD_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
            DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
            DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", UUpdated, DBUTIL.FieldTypes.ftText)
            SQL = DB.CombineSQL(op, SQL1, SQL2, "SYS_USERS", "USER_NAME=" & DB.SQLValue(UID, DBUTIL.FieldTypes.ftText).ToString().Trim())
            result = DB.ExecSQL(SQL)

            Return (result > 0)
        Catch ex As Exception
            Throw New DALException(ex)
        End Try
    End Function

    Public Function SearchRoleList(Optional ByVal RoleID As String = "", Optional ByVal RoleName As String = "" _
    , Optional ByVal RoleDesc As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim DT As New DataTable
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(ROLE_NAME)", UCase(RoleName), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(ROLE_DESC)", UCase(RoleDesc), DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM SYS_ROLES "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY ROLE_ID"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try

    End Function

    ''Updated By Aoy 23/06/2552 
    'Public Function SearchUserList(ByVal UserName As String, Optional ByVal UserDesc As String = "" _
    ', Optional ByVal RoleID As String = "", Optional ByVal UserLevel As String = "", Optional ByVal UserGroup As String = "" _
    ', Optional ByVal Status As String = "", Optional ByVal UserType As String = "" _
    ', Optional ByVal UserCode As String = "", Optional ByVal ExpireDateF As String = "" _
    ', Optional ByVal ExpireDateT As String = "", Optional ByVal Email As String = "" _
    ', Optional ByVal TelNo As String = "", Optional ByVal Mobile As String = "" _
    ', Optional ByVal OtherCriteria As String = "") As DataTable
    '    Dim DT As New DataTable
    '    Dim SQL As String = "", Criteria As String = ""

    '    Try
    '        Criteria = OtherCriteria
    '        DB.AddCriteria(Criteria, "UPPER(SU.USER_NAME)", UserName.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SU.USER_DESC)", UserDesc.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SU.ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SU.USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SU.GROUP_ID", UserGroup, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SU.DISABLED_FLAG", Status, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SU.USER_TYPE", UserType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "DECODE(SU.DAY_EXPIRE,0,NULL,(SU.CHG_PWD_DATE + SU.DAY_EXPIRE))", AppDateValue(ExpireDateF), AppDateValue(ExpireDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria, "UPPER(SU.USER_EMAIL)", Email.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SU.TEL_NO)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SU.MOBILE_NO)", Mobile.ToUpper, DBUTIL.FieldTypes.ftText)

    '        SQL = "SELECT SU.*, SU.CHG_PWD_DATE + SU.DAY_EXPIRE AS PWD_EXPIRE_DATE" & _
    '        ", SR.ROLE_NAME,SR.ROLE_DESC,UL.LEVEL_NAME,UG.RIGHTS,UG.GROUP_NAME,UG.PERMIS_INFOS,UG.PERMIS_PROJECT_TYPES,UG.PERMIS_HIST " & _
    '        ",DECODE(SU.DISABLED_FLAG,'Y','Enabled','Disabled') AS DISABLE_FLAG_DESC " & _
    '        "FROM SYS_USERS SU, SYS_ROLES SR, REF_USER_LEVELS UL, SYS_GROUPS UG WHERE " & _
    '        "SU.ROLE_ID = SR.ROLE_ID(+) AND SU.USER_LEVEL=UL.LEVEL_ID(+) AND SU.GROUP_ID=UG.GROUP_ID(+)"
    '        If Criteria <> "" Then SQL &= " AND " & Criteria
    '        SQL &= " ORDER BY SU.ROLE_ID, SU.USER_NAME"

    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception
    '        Throw New DALException(ex)
    '    Finally
    '        ClearObject(DT)
    '    End Try
    'End Function

    Public Function SearchUserList(ByVal UserName As String, Optional ByVal UserDesc As String = "" _
    , Optional ByVal RoleID As String = "", Optional ByVal UserLevel As String = "", Optional ByVal UserGroup As String = "" _
    , Optional ByVal Status As String = "", Optional ByVal UserType As String = "" _
    , Optional ByVal UserCode As String = "", Optional ByVal ExpireDateF As String = "" _
    , Optional ByVal ExpireDateT As String = "", Optional ByVal Email As String = "" _
    , Optional ByVal TelNo As String = "", Optional ByVal Mobile As String = "" _
    , Optional ByVal OtherCriteria As String = "") As DataTable
        Dim DT As New DataTable
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(USER_NAME)", UserName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(USER_DESC)", UserDesc.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "GROUP_ID", UserGroup, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "DISABLED_FLAG", Status, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "USER_TYPE", UserType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "DECODE(DAY_EXPIRE,0,NULL,PWD_EXPIRE_DATE)", AppDateValue(ExpireDateF), AppDateValue(ExpireDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(USER_EMAIL)", Email.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(TEL_NO)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(MOBILE_NO)", Mobile.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * " & _
            "FROM V_SYS_USERS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY ROLE_ID, USER_NAME"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    'Aoy 09/12/51
    Public Function SearchRefUnitView() As DataTable
        Dim DT As New DataTable
        Dim SQL As String = "", Criteria As String = ""

        Try
            'DB.AddCriteria(Criteria, "SU.USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "SU.USER_DESC", UserDesc, DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "SU.ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM REF_UNIT_VIEW "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY UNIT_VIEW "

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    'Updated By Aoy 04/05/2552
    Public Function MngUserData(ByVal op As Integer, ByRef UserName As String, _
    Optional ByVal UserDesc As String = Nothing, Optional ByVal Password As String = Nothing, _
    Optional ByVal Role As String = Nothing, Optional ByVal Disable As String = Nothing, _
    Optional ByVal EmpCode As String = Nothing, Optional ByVal Exp As String = Nothing, _
    Optional ByVal TelNo As String = Nothing, Optional ByVal Mobile As String = Nothing, _
    Optional ByVal Email As String = Nothing, Optional ByVal UserLevel As String = Nothing, _
    Optional ByVal UserGroup As String = Nothing, Optional ByVal UserType As String = Nothing, _
    Optional ByVal Code As String = Nothing, Optional ByVal BannedFlag As String = Nothing, _
    Optional ByVal ClearChangePWDDate As Boolean = False) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT

                    DB.AddSQL(op, SQL1, SQL2, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                    DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
                Else
                    op = DBUTIL.opUPDATE

                    DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                    DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
                End If

                'If Not IsNothing(UserDesc) Then DB.AddSQL(op, SQL1, SQL2, "USER_DESC", UserDesc, DBUTIL.FieldTypes.ftText)

                If Not IsNothing(Password) Then
                    DB.AddSQL(op, SQL1, SQL2, "PASSWORD", Password, DBUTIL.FieldTypes.ftText)
                    If ClearChangePWDDate Then
                        DB.AddSQL(op, SQL1, SQL2, "CHG_PWD_DATE", "", DBUTIL.FieldTypes.ftDate)
                    Else
                        DB.AddSQL(op, SQL1, SQL2, "CHG_PWD_DATE", Now, DBUTIL.FieldTypes.ftDate)
                    End If
                End If

                DB.AddSQL2(op, SQL1, SQL2, "ROLE_ID", Role, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "USER_DESC", UserDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DISABLED_FLAG", Disable, DBUTIL.FieldTypes.ftText)
                'DB.AddSQL2(op, SQL1, SQL2, "EMP_CODE", EmpCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "MOBILE_NO", Mobile, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "USER_EMAIL", Email, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DAY_EXPIRE", Exp, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "GROUP_ID", UserGroup, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "USER_TYPE", UserType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "CODE", Code, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BANNED_FLAG", BannedFlag, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "SYS_USERS", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Send Mail 26/03/2552 
    Public Function SearchTemplate(Optional ByVal MailTempID As String = "") As DataTable   'ดูจากตาราง MAIL_TEMPLATES เพื่อเอาข้อมูลของ content/template มาดูนะ
        Dim DT As New DataTable
        Dim SQL As String = "", Criteria As String = ""
        Try
            DB.AddCriteria(Criteria, "MAIL_TEMP_ID", MailTempID, DBUTIL.FieldTypes.ftNumeric)
            ' DB.AddCriteria(Criteria, "CONTENT_NAME", contHeader, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT *  FROM MAIL_TEMPLATES"

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY MAIL_TEMP_ID"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    Public Function SearchTaskList() As DataTable
        Dim DT As New DataTable
        Dim SQL As String

        Try
            SQL = "SELECT * FROM SYS_TASKS ORDER BY TASK_ID"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    Public Function MngRoleData(ByVal op As Integer, ByRef RoleID As String, Optional ByVal RoleName As String = Nothing, Optional ByVal RoleDesc As String = Nothing, Optional ByVal RoleRights As String = Nothing, Optional ByVal User As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    RoleID = GenerateID("SYS_ROLES", "ROLE_ID")
                    DB.AddSQL(op, SQL1, SQL2, "ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opUPDATE
                End If

                If Not IsNothing(RoleName) Then DB.AddSQL(op, SQL1, SQL2, "ROLE_NAME", RoleName, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(RoleDesc) Then DB.AddSQL(op, SQL1, SQL2, "ROLE_DESC", RoleDesc, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(RoleRights) Then DB.AddSQL(op, SQL1, SQL2, "RIGHTS", RoleRights, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "SYS_ROLES", Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngGroupData(ByVal op As Integer, ByRef GroupID As String, Optional ByVal GroupName As String = Nothing _
    , Optional ByVal RoleID As String = Nothing, Optional ByVal GroupRights As String = Nothing, Optional ByVal PermisInfos As String = Nothing _
    , Optional ByVal PermisHist As String = Nothing, Optional ByVal PermisProjectTypes As String = Nothing _
    , Optional ByVal PermisSVResponse As String = Nothing, Optional ByVal GroupType As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "GROUP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    GroupID = GenerateID("SYS_GROUPS", "GROUP_ID")
                    DB.AddSQL(op, SQL1, SQL2, "GROUP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opUPDATE
                End If
                DB.AddSQL2(op, SQL1, SQL2, "ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "GROUP_NAME", GroupName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "RIGHTS", GroupRights, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PERMIS_INFOS", PermisInfos, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PERMIS_HIST", PermisHist, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PERMIS_PROJECT_TYPES", PermisProjectTypes, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PERMIS_SV_RESPONSE", PermisSVResponse, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "GROUP_TYPE", GroupType, DBUTIL.FieldTypes.ftNumeric)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "SYS_GROUPS", Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SaveUserData(ByVal UID As String, ByVal UDesc As String, ByVal UMail As String, ByVal UPwd As String, _
                                 ByVal URole As Integer, ByVal UExp As Integer, ByVal UDisable As String, ByVal UApprover As String) As String

        Dim DT As DataTable = Nothing
        Dim DR As DataRow
        Dim SQL1, SQL2, SQL As String
        Dim op As Integer

        Try
            UID = UID.ToUpper()
            SQL = "" : SQL1 = "" : SQL2 = ""

            DT = SearchUserList(UID, "", "")
            DR = GetDR(DT)
            If Not IsNothing(DR) Then
                op = DBUTIL.opUPDATE
            Else
                op = DBUTIL.opINSERT
                DB.AddSQL(op, SQL1, SQL2, "USER_NAME", UID, DBUTIL.FieldTypes.ftText)
            End If

            DB.AddSQL(op, SQL1, SQL2, "USER_DESC", UDesc, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op, SQL1, SQL2, "USER_EMAIL", UMail, DBUTIL.FieldTypes.ftText)
            If UPwd <> "****" Then DB.AddSQL(op, SQL1, SQL2, "PASSWORD", Encrypted(UID, UPwd), DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op, SQL1, SQL2, "ROLE_ID", URole, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(op, SQL1, SQL2, "DAY_EXPIRE", UExp, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddSQL(op, SQL1, SQL2, "APPROVER_CODE", UApprover, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op, SQL1, SQL2, "DISABLED_FLAG", UDisable, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op, SQL1, SQL2, "CHG_PWD_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
            DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
            DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME"), DBUTIL.FieldTypes.ftText)

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SYS_USERS", "USER_NAME=" & DB.SQLValue(UID, DBUTIL.FieldTypes.ftText).ToString().Trim())
            DB.ExecSQL(SQL)

            Return ""
        Catch ex As Exception
            Throw New DALException(ex)
        Finally
            ClearObject(DT)
        End Try
    End Function

    Public Function DeleteUserData(ByVal UserName As String) As String
        Dim SQL As String

        Try
            SQL = "DELETE FROM SYS_USERS WHERE USER_NAME = '" & UserName & "'"

            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw New DALException(ex.Message)
        End Try
    End Function

    'Updated By Aoy 23/06/2552
    Public Function Login(ByVal UserName As String, Optional ByVal Password As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", CriteriaSQL As String = ""
        Dim status As String = ""

        If UserName <> "" Then
            DB.AddCriteria(CriteriaSQL, "SU.USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SU.PASSWORD", Password, DBUTIL.FieldTypes.ftText)

            'SQL = "SELECT SU.*, SU.CHG_PWD_DATE + SU.DAY_EXPIRE AS PWD_EXPIRE_DATE, SR.ROLE_DESC AS ROLE_DESC, SR.RIGHTS AS RIGHTS, DBU.BU_CODE, DBU.BU_DESC " & _
            '      "FROM SYS_USERS SU ,SYS_ROLES SR, SYS_USERS AP, DIM_BUSINESS_UNITS DBU " & _
            '      "WHERE SU.COMPANY_ID *= DBU.BU_ID AND SU.ROLE_ID = SR.ROLE_ID"

            'pui 14/5/52
            SQL = "SELECT SU.*, SU.CHG_PWD_DATE + SU.DAY_EXPIRE AS PWD_EXPIRE_DATE, SR.ROLE_DESC AS ROLE_DESC, SG.RIGHTS AS RIGHTS" & _
            ",SG.GROUP_NAME,SG.PERMIS_INFOS,SG.PERMIS_PROJECT_TYPES,SG.PERMIS_HIST,SG.PERMIS_SV_RESPONSE  " & _
                  "FROM SYS_USERS SU,SYS_ROLES SR,SYS_GROUPS SG WHERE " & _
                  "SU.ROLE_ID = SR.ROLE_ID(+) AND SU.GROUP_ID=SG.GROUP_ID(+)"

            If CriteriaSQL <> "" Then SQL &= " AND " & CriteriaSQL
            DB.OpenDT(DT, SQL)
        End If

        Return DT
    End Function
#End Region

#Region "Lookup"
    Public Function SearchConfigs(Optional ByVal CfgTable As String = "", Optional ByVal CfgName As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "CFG_TABLE", CfgTable, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "CFG_NAME", CfgName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT *  FROM CONFIGS "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY CFG_NAME"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function SearchConfigLookup(Optional ByVal CfgCode As String = "", Optional ByVal CfgName As String = "", Optional ByVal ParentID1 As String = "", Optional ByVal ParentID2 As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim SQL As String = "", Criteria As String = ""
        Dim CfgGroup As Integer = 0, CfgTable As String = "", KeyIDName As String = "", KeyValueName As String = ""
        Dim KeyValueName2 As String = "NULL"
        Dim KeyValueName3 As String = "NULL"
        Dim KeyValueName4 As String = "NULL"
        Dim ParentKeyID1 As String = "NULL"
        Dim ParentKeyID2 As String = "NULL"
        Dim ParentKeyType1 As String = "NULL"
        Dim ParentKeyType2 As String = "NULL"
        Dim FType As Integer = 2
        Dim CriteriaSQL As String = ""
        Dim ActiveFlagField As String = "NULL"
        Dim ReadOnlyFlag As String = "NULL"
        Try
            DT = SearchConfigs(CfgCode)
            DR = DT.Rows(0)
            CfgGroup = CInt("0" & DR("CFG_GROUP") & "")
            CfgTable = DR("CFG_TABLE") & ""
            KeyIDName = DR("CFG_KEY_ID") & ""
            KeyValueName = DR("CFG_KEY_NAME1") & ""
            If DR("CFG_KEY_NAME2") & "" <> "" Then KeyValueName2 = DR("CFG_KEY_NAME2") & ""
            If DR("CFG_KEY_NAME3") & "" <> "" Then KeyValueName3 = DR("CFG_KEY_NAME3") & ""
            If DR("CFG_KEY_NAME4") & "" <> "" Then KeyValueName4 = DR("CFG_KEY_NAME4") & ""
            If DR("CFG_PARENT_ID1") & "" <> "" Then ParentKeyID1 = DR("CFG_PARENT_ID1") & ""
            If DR("CFG_PARENT_ID2") & "" <> "" Then ParentKeyID2 = DR("CFG_PARENT_ID2") & ""
            If DR("CFG_PARENT_TYPE1") & "" <> "" Then ParentKeyType1 = DR("CFG_PARENT_TYPE1") & ""
            If DR("CFG_PARENT_TYPE2") & "" <> "" Then ParentKeyType2 = DR("CFG_PARENT_TYPE2") & ""
            If DR("CFG_ACTIVE") & "" <> "" Then ActiveFlagField = DR("CFG_ACTIVE") & ""
            If DR("CFG_READONLY") & "" <> "" Then ReadOnlyFlag = "'" & DR("CFG_READONLY") & "'"

            If CfgGroup > 0 Then
                DT = SearchLookupData(CfgGroup.ToString(), LevelID:="1")
                If Not IsNothing(DT) Then
                    DT.Columns.Add("LK_VALUE3", System.Type.GetType("System.String"))
                    DT.Columns.Add("LK_VALUE4", System.Type.GetType("System.String"))
                    DT.Columns.Add("PARENT_ID1", System.Type.GetType("System.String"))
                    DT.Columns.Add("PARENT_ID2", System.Type.GetType("System.String"))
                    Dim DC As New DataColumn
                    DC.ColumnName = "COL_CNT"
                    DC.DataType = System.Type.GetType("System.Int16")
                    DC.DefaultValue = "1"
                    DT.Columns.Add(DC)
                    DT.AcceptChanges()
                End If
            Else
                SQL = "SELECT 0 AS LK_GRP_ID," & KeyIDName & " AS LK_ITEM_ID," & KeyValueName & " AS LK_VALUE," & _
                                KeyValueName2 & " AS LK_VALUE2, " & KeyValueName3 & " AS LK_VALUE3, " & _
                                KeyValueName4 & " AS LK_VALUE4, " & ParentKeyID1 & " AS PARENT_ID1, " & _
                                ParentKeyID2 & " AS PARENT_ID2, " & ActiveFlagField & " AS ACTIVE_FLAG, " & _
                                ReadOnlyFlag & " AS READ_FLAG " & _
                                " FROM " & CfgTable
                If ParentID1 <> "" Then
                    If ParentKeyType1 = "TEXT" Then
                        CriteriaSQL = ParentKeyID1 & " LIKE '" & ParentID1 & "%'"
                    Else
                        CriteriaSQL = ParentKeyID1 & " = " & ParentID1 & ""
                    End If
                End If
                If ParentID2 <> "" Then
                    If CriteriaSQL <> "" Then CriteriaSQL += "AND "
                    If ParentKeyType2 = "TEXT" Then
                        If ParentID2 = "NULL" Then
                            CriteriaSQL += ParentKeyID2 & " IS NULL "
                        Else
                            CriteriaSQL += ParentKeyID2 & " LIKE '" & ParentID2 & "%'"
                        End If
                    Else
                        CriteriaSQL += ParentKeyID2 & " = " & ParentID2 & ""
                    End If
                End If

                If CriteriaSQL <> "" Then
                    SQL += " WHERE " + CriteriaSQL
                End If
                SQL += " ORDER BY " & KeyIDName

                DT = QueryData(SQL)
                If Not IsNothing(DT) Then
                    Dim DC As New DataColumn
                    DC.ColumnName = "COL_CNT"
                    DC.DataType = System.Type.GetType("System.Int16")
                    DC.DefaultValue = "1"
                    If KeyValueName2 <> "NULL" Then
                        DC.DefaultValue = "2"
                    ElseIf KeyValueName3 <> "NULL" Then
                        DC.DefaultValue = "3"
                    End If
                    DT.Columns.Add(DC)
                    DT.AcceptChanges()
                End If
            End If
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngConfigLookup(ByVal op As Integer, Optional ByVal CfgCode As String = "", Optional ByVal ItemID As String = "", Optional ByVal ItemIDOld As String = "", Optional ByVal Desc As String = "", Optional ByVal Desc2 As String = "", _
                                    Optional ByVal Desc3 As String = "", Optional ByVal Desc4 As String = "", Optional ByVal ParentID1 As String = "", Optional ByVal ParentID2 As String = "", Optional ByVal User As String = "") As String
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim CfgGroup As Integer = 0, CfgTable As String = "", KeyIDName As String = "", KeyValueName As String = "", KeyIDType As String = ""
        Dim FType As Integer = 2
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""
        Dim KeyValueName2 As String = ""
        Dim KeyValueName3 As String = ""
        Dim KeyValueName4 As String = ""
        Dim ParentKeyID1 As String = ""
        Dim ParentKeyID2 As String = ""
        Dim ParentKeyType1 As String = ""
        Dim ParentKeyType2 As String = ""
        Dim CriteriaSQL As String = ""
        Dim ParentType1 As Integer = 2
        Dim ParentType2 As Integer = 2
        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            DT = SearchConfigs(CfgCode)
            DR = DT.Rows(0)
            CfgGroup = CInt("0" & DR("CFG_GROUP") & "")
            KeyIDName = DR("CFG_KEY_ID") & ""
            KeyValueName = DR("CFG_KEY_NAME1") & ""
            KeyIDType = DR("CFG_KEY_TYPE") & ""
            CfgTable = DR("CFG_TABLE") & ""
            If DR("CFG_KEY_NAME2") & "" <> "" Then KeyValueName2 = DR("CFG_KEY_NAME2") & ""
            If DR("CFG_KEY_NAME3") & "" <> "" Then KeyValueName3 = DR("CFG_KEY_NAME3") & ""
            If DR("CFG_KEY_NAME4") & "" <> "" Then KeyValueName4 = DR("CFG_KEY_NAME4") & ""
            If DR("CFG_PARENT_ID1") & "" <> "" Then ParentKeyID1 = DR("CFG_PARENT_ID1") & ""
            If DR("CFG_PARENT_ID2") & "" <> "" Then ParentKeyID2 = DR("CFG_PARENT_ID2") & ""
            If DR("CFG_PARENT_TYPE1") & "" <> "" Then ParentKeyType1 = DR("CFG_PARENT_TYPE1") & ""
            If DR("CFG_PARENT_TYPE2") & "" <> "" Then ParentKeyType2 = DR("CFG_PARENT_TYPE2") & ""

            Select Case KeyIDType
                Case "NUMERIC" : FType = DBUTIL.FieldTypes.ftNumeric
                Case "TEXT" : FType = DBUTIL.FieldTypes.ftText
            End Select

            If ParentKeyType1 <> "" Then
                Select Case ParentKeyType1
                    Case "NUMERIC" : ParentType1 = DBUTIL.FieldTypes.ftNumeric
                    Case "TEXT" : ParentType1 = DBUTIL.FieldTypes.ftText
                End Select
            End If

            If ParentKeyType2 <> "" Then
                Select Case ParentKeyType2
                    Case "NUMERIC" : ParentType2 = DBUTIL.FieldTypes.ftNumeric
                    Case "TEXT" : ParentType2 = DBUTIL.FieldTypes.ftText
                End Select
            End If


            If CfgGroup > 0 Then
                Return MngLookupData(op, CfgGroup.ToString(), ItemID, ItemIDOld, Desc, Desc2, "0", "1", User)
            Else
                If op <> DBUTIL.opINSERT Then
                    DB.AddCriteria(Criteria, KeyIDName, ItemIDOld, FType)

                End If
                If op <> DBUTIL.opDELETE Then
                    If op = DBUTIL.opINSERT Then
                        op = DBUTIL.opINSERT
                        DB.AddSQL(op, SQL1, SQL2, KeyIDName, ItemID, FType)

                        DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                        DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", User, DBUTIL.FieldTypes.ftText)
                    Else
                        op = DBUTIL.opUPDATE
                        DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                        DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", User, DBUTIL.FieldTypes.ftText)
                    End If

                    If Not IsNothing(Desc) Then DB.AddSQL(op, SQL1, SQL2, KeyValueName, Desc, DBUTIL.FieldTypes.ftText)
                    If Not IsNothing(Desc2) Then DB.AddSQL(op, SQL1, SQL2, KeyValueName2, Desc2, DBUTIL.FieldTypes.ftText)
                    If Not IsNothing(Desc3) Then DB.AddSQL(op, SQL1, SQL2, KeyValueName3, Desc3, DBUTIL.FieldTypes.ftText)
                    If Not IsNothing(Desc4) Then DB.AddSQL(op, SQL1, SQL2, KeyValueName4, Desc4, DBUTIL.FieldTypes.ftText)
                    If Not IsNothing(ParentID1) Then DB.AddSQL(op, SQL1, SQL2, ParentKeyID1, ParentID1, ParentType1)
                    If Not IsNothing(ParentID2) Then DB.AddSQL(op, SQL1, SQL2, ParentKeyID2, ParentID2, ParentType2)
                End If

                If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                    Throw New Exception("Insufficient data!")
                Else
                    SQL = DB.CombineSQL(op, SQL1, SQL2, CfgTable, Criteria)
                    DB.ExecSQL(SQL)
                End If

                Return ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchLookupGroup(Optional ByVal GroupID As String = "", Optional ByVal GroupCode As String = "", Optional ByVal GroupDesc As String = "", Optional ByVal EditFlag As String = "Y", Optional ByVal ActiveFlag As String = "Y") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "LK_GRP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "LK_GRP_CODE", GroupCode, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "LK_GRP_DESC", GroupDesc, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EDITABLE_FLAG", EditFlag, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "ACTIVE_FLAG", ActiveFlag, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM LOOKUP_GROUPS "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY LK_GRP_ID"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchLookupData(Optional ByVal GroupID As String = "", Optional ByVal GroupCode As String = "", Optional ByVal Key As String = "", Optional ByVal Desc As String = "", Optional ByVal ParentID As String = "", Optional ByVal LevelID As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "D.LK_GRP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "D.LK_ITEM_ID", Key, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "D.LK_VALUE", Desc, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "D.PARENT_ITEM_ID", ParentID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "D.ITEM_LEVEL", LevelID, DBUTIL.FieldTypes.ftNumeric)

            If GroupCode = "" Then
                SQL = "SELECT D.* FROM LOOKUP_DATA D "
                If Criteria <> "" Then SQL &= " WHERE " & Criteria
            Else
                DB.AddCriteria(Criteria, "G.LK_GRP_CODE", GroupCode, DBUTIL.FieldTypes.ftText)

                SQL = "SELECT D.* FROM LOOKUP_DATA D, LOOKUP_GROUPS G WHERE D.LK_GRP_ID = G.LK_GRP_ID "
                If Criteria <> "" Then SQL &= " AND " & Criteria
            End If
            SQL &= " ORDER BY D.LK_GRP_ID, D.LK_ITEM_ID"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngLookupData(ByVal op As Integer, ByVal GroupID As String, ByRef Key As String, ByRef KeyOld As String, Optional ByVal Desc As String = Nothing, Optional ByVal Desc2 As String = Nothing, Optional ByVal ParentID As String = Nothing, Optional ByVal Level As String = Nothing, Optional ByVal User As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "LK_GRP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "LK_ITEM_ID", KeyOld, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "ITEM_LEVEL", Level, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "PARENT_ITEM_ID", ParentID, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT

                    'Gen New ID
                    'SQL = "SELECT MAX(LK_ITEM_ID) AS MAX FROM LOOKUP_DATA WHERE LK_GRP_ID = " & GroupID
                    'Key = DB.LookupSQL(SQL)
                    'Key = Val(Key) + 1
                    'SQL = ""

                    DB.AddSQL(op, SQL1, SQL2, "LK_GRP_ID", GroupID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "LK_ITEM_ID", Key, DBUTIL.FieldTypes.ftText)
                    If Not IsNothing(Level) Then DB.AddSQL(op, SQL1, SQL2, "ITEM_LEVEL", Level, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(ParentID) Then DB.AddSQL(op, SQL1, SQL2, "PARENT_ITEM_ID", ParentID, DBUTIL.FieldTypes.ftText)

                    DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                    DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", User, DBUTIL.FieldTypes.ftText)
                Else
                    op = DBUTIL.opUPDATE

                    DB.AddSQL(op, SQL1, SQL2, "LK_ITEM_ID", Key, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                    DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", User, DBUTIL.FieldTypes.ftText)
                End If

                If Not IsNothing(Desc) Then DB.AddSQL(op, SQL1, SQL2, "LK_VALUE", Desc, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(Desc2) Then DB.AddSQL(op, SQL1, SQL2, "LK_VALUE2", Desc2, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "LOOKUP_DATA", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Audit"
    'Public Sub InsertAudit(ByVal transaction As PTT.INTERNET.AuditData)
    '    Dim SQL1, SQL2, SQL As String
    '    'Dim newTRANS_ID As Integer

    '    Try
    '        SQL1 = "" : SQL2 = ""
    '        'DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "TRANS_ID", "SYS_LOG_SEQ.nextval", 4)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "TRANS_DATE", transaction.TransactionDate, DBUTIL.FieldTypes.ftDateTime)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "CATEGORY", transaction.Category & "", DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "ACTION_DETAIL", Left(Trim(transaction.Action & ""), 250), DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "USER_ID", transaction.UserId & "", DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "DEPARTMENT_NAME", transaction.DeptName & "", DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "MACHINE_NAME", transaction.MachineName & "", DBUTIL.FieldTypes.ftText)

    '        SQL = DB.CombineSQL(DBUTIL.opINSERT, SQL1, SQL2, "SYS_LOGS", "")
    '        DB.ExecSQL(SQL)


    '    Catch ex As Exception

    '        Throw New DALException(ex.Message)
    '    End Try
    'End Sub

    'Public Function SearchAudit(ByVal TransactionFromDate As Date, ByVal TransactionToDate As Date, _
    '                            ByVal UserId As String, ByVal Deptname As String, ByVal Category As String, ByVal Action As String) As PTT.INTERNET.AuditDatas
    '    Dim results As New PTT.INTERNET.AuditDatas
    '    Dim DT As DataTable
    '    Dim DR As DataRow
    '    Dim SQL As String
    '    Dim CriteriaSQL As String = ""
    '    Dim CategorySQL As String = ""
    '    'Dim Conn As OleDb.OleDbConnection

    '    Try
    '        'If UseStoredProc Then
    '        '    DB.InitParams()

    '        '    DB.AddParam("@TRANS_FROMDATE", TransactionFromDate, FieldTypes.ftDate)
    '        '    DB.AddParam("@TRANS_TODATE", TransactionToDate, FieldTypes.ftDate)
    '        '    DB.AddParam("@USER_ID", Replace(UserId, "*", "%") & "", FieldTypes.ftText)
    '        '    DB.AddParam("@CATEGORY", Replace(Category, "*", "%") & "", FieldTypes.ftText)
    '        '    DB.AddParam("@ACTION", Replace(Action, "*", "%") & "", FieldTypes.ftText)


    '        '    DT = GetDT(DB.ExecProcDS("SearchAudit"))
    '        'Else
    '        DB.AddCriteriaRange(CriteriaSQL, "TRANS_DATE", TransactionFromDate, TransactionToDate, DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(CriteriaSQL, "UPPER(ACTION_DETAIL)", Action, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(CriteriaSQL, "UPPER(USER_ID)", UserId, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(CriteriaSQL, "DEPARTMENT_NAME", Deptname, DBUTIL.FieldTypes.ftText)
    '        If Category = "ERROR" Then
    '            DB.AddCriteria(CriteriaSQL, "CATEGORY", Category, DBUTIL.FieldTypes.ftText)
    '        Else
    '            CategorySQL = " (CATEGORY = 'LOG' OR CATEGORY='MAIL' ) "
    '            If CriteriaSQL <> "" Then
    '                CriteriaSQL += " AND  " + CategorySQL
    '            Else
    '                CriteriaSQL = CategorySQL
    '            End If
    '        End If

    '        SQL = "SELECT * " & _
    '              "FROM SYS_LOGS "
    '        If CriteriaSQL <> "" Then SQL += "WHERE " + CriteriaSQL
    '        SQL += " ORDER BY TRANS_ID DESC"

    '        DT = Nothing
    '        DB.OpenDT(DT, SQL)
    '        'End If


    '        For Each DR In DT.Rows
    '            results.Add(New PTT.INTERNET.AuditData(DR("TRANS_DATE"), DR("USER_ID") & "", DR("DEPARTMENT_NAME") & "", _
    '                            DR("CATEGORY") & "", DR("ACTION_DETAIL") & "", _
    '                            DR("MACHINE_NAME") & ""))
    '        Next
    '        DB.CloseDT(DT)

    '    Catch ex As Exception
    '        results = Nothing
    '        Throw New DALException(ex.Message)
    '    End Try

    '    Return results
    'End Function

    Public Sub InsertAudit(ByVal Category As String, ByVal Action As String, ByVal User As String, ByVal IPAddress As String _
    , Optional ByVal RefID1 As String = "", Optional ByVal RefID2 As String = "" _
    , Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing)
        'Dim Conn As OleDb.OleDbConnection = Nothing
        'Dim Trans As OleDb.OleDbTransaction = Nothing
        Dim SQL1, SQL2, SQL As String
        Dim NewID As Integer

        Try
            'Conn = DB.OpenConn(_ConnectStr)
            'Trans = DB.BeginTrans(Conn)

            NewID = GenerateID("SYS_LOGS", "TRANS_ID", Conn:=Conn, Trans:=Trans)

            SQL1 = "" : SQL2 = "" : SQL = ""
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "TRANS_ID", NewID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "TRANS_DATE", System.DateTime.Now, DBUTIL.FieldTypes.ftDateTime)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "CATEGORY", Category, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "ACTION_DETAIL", Left(Trim(Action), 500), DBUTIL.FieldTypes.ftText)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "USER_NAME", User, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "IP_ADDRESS", IPAddress, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "REF_ID1", RefID1, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(DBUTIL.opINSERT, SQL1, SQL2, "REF_ID2", RefID2, DBUTIL.FieldTypes.ftNumeric)

            SQL = DB.CombineSQL(DBUTIL.opINSERT, SQL1, SQL2, "SYS_LOGS", "")
            DB.ExecSQL(SQL, Conn, Trans)
        Catch ex As Exception
            Throw New DALException(ex.Message)
        End Try
    End Sub

    Public Function SearchAudit(ByVal FromDate As String, ByVal ToDate As String, ByVal Category As String _
    , ByVal Action As String, ByVal UserName As String, Optional ByVal UserDesc As String = "" _
    , Optional ByVal UserLevel As String = "", Optional ByVal RoleID As String = "" _
    , Optional ByVal TaskID As String = "", Optional ByVal Email As String = "" _
    , Optional ByVal TelNo As String = "", Optional ByVal MobileNo As String = "" _
    , Optional ByVal UserStatus As String = "", Optional ByVal IPAddress As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderSQL As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String
        Dim CriteriaSQL As String = ""
        Dim CategorySQL As String = ""

        Try
            CriteriaSQL = OtherCriteria
            DB.AddCriteriaRange(CriteriaSQL, "SL.TRANS_DATE", AppDateValue(FromDate), AppDateValue(ToDate), DBUTIL.FieldTypes.ftDateTime)
            DB.AddCriteria(CriteriaSQL, "SL.CATEGORY", Category, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SL.ACTION_DETAIL", Action, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "UPPER(SL.USER_NAME)", UserName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "UPPER(SU.USER_DESC)", UserDesc.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SU.USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(CriteriaSQL, "SU.ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(CriteriaSQL, "SL.REF_ID1", TaskID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(CriteriaSQL, "UPPER(SU.USER_EMAIL)", Email.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SU.TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SU.MOBILE_NO", MobileNo, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SL.IP_ADDRESS", IPAddress, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(CriteriaSQL, "SU.DISABLED_FLAG", UserStatus, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT SL.*,SU.USER_DESC,SU.ROLE_ID,SU.USER_EMAIL,SU.TEL_NO,SU.MOBILE_NO,SU.USER_LEVEL,SR.ROLE_NAME,ST.TASK_DESC " & _
            " ,DECODE(SU.DISABLED_FLAG,'N','Enable','Disable') As USER_STATUS,UL.LEVEL_NAME " & _
            " FROM SYS_LOGS SL,SYS_USERS SU,SYS_ROLES SR,SYS_TASKS ST,REF_USER_LEVELS UL " & _
            "WHERE SL.USER_NAME=SU.USER_NAME(+) AND SU.ROLE_ID=SR.ROLE_ID(+) " & _
            " AND SL.REF_ID1=ST.TASK_ID(+) AND SU.USER_LEVEL=UL.LEVEL_ID(+)"
            If CriteriaSQL <> "" Then SQL &= "AND " & CriteriaSQL
            If OrderSQL = "" Then
                SQL &= " ORDER BY SL.TRANS_ID DESC"
            Else
                SQL &= " ORDER BY " & OrderSQL
            End If

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw New DALException(ex.Message)
        Finally
            ClearObject(DT)
        End Try
    End Function
#End Region

#Region "MasterData"
    Public Function SearchSiteType(Optional ByVal SiteType As String = "", Optional ByVal SiteTypeDesc As String = "", _
                               Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SITE_TYPE", SiteType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SITE_TYPE_DESC", SiteTypeDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_SITE_TYPES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SITE_TYPE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchNetworkType(Optional ByVal NetworkType As String = "", Optional ByVal NetworkTypeDesc As String = "", _
                                       Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "NETWORK_TYPE_DESC", NetworkTypeDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_NETWORK_TYPES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY NETWORK_TYPE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchOwnerType(Optional ByVal OwnerType As String = "", Optional ByVal OwnerTypeDesc As String = "", _
                                       Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "OWNER_TYPE", OwnerType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "OWNER_TYPE_DESC", OwnerTypeDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_OWNER_TYPES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY OWNER_TYPE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchNetworkProtocal(Optional ByVal NetProtocol As String = "", Optional ByVal NetProtocolDesc As String = "", _
                                   Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "NET_PROTOCOL", NetProtocol, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "NET_PROTOCOL_DESC", NetProtocolDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_NET_PROTOCOLS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY NET_PROTOCOL"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchBrand(Optional ByVal BrandID As String = "", Optional ByVal BrandName As String = "", _
                               Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "BRAND_ID", BrandID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "BRAND_NAME", BrandName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_BRANDS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY BRAND_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchModel(Optional ByVal ModelID As String = "", Optional ByVal ModelName As String = "", Optional ByVal BandID As String = "", _
                               Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "MODEL_ID", ModelID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "MODEL_NAME", ModelName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "BRAND_ID", BandID, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_MODELS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY MODEL_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSaleArea(Optional ByVal SaleArea As String = "", Optional ByVal SaleAreaName As String = "", _
                                   Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SALE_AREA", SaleArea, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SALE_AREA_NAME", SaleAreaName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_SALE_AREAS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SALE_AREA"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchPlant(Optional ByVal PlantCode As String = "", Optional ByVal PlantENDesc As String = "", Optional ByVal PlantTHDesc As String = "", _
                               Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PLANT_CODE", PlantCode, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PLANT_EN_NAME", PlantENDesc, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PLANT_TH_NAME", PlantENDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM SAP_PLANTS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PLANT_CODE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchBranch(Optional ByVal BranchCode As String = "", Optional ByVal BranchName As String = "", _
                              Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "BRANCH_CODE", BranchCode, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "BRANCH_NAME", BranchName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_BRANCHES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY BRANCH_CODE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchProjectType(Optional ByVal ProjectType As String = "", Optional ByVal ProjectTypeDesc As String = "", _
                                 Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PROJECT_TYPE_DESC", ProjectTypeDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_PROJECT_TYPES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PROJECT_TYPE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function



    Public Function SearchVendor(Optional ByVal VendorCode As String = "", Optional ByVal VendorName As String = "", _
                                 Optional ByVal StaffName As String = "", Optional ByVal ContactName As String = "", _
                                 Optional ByVal VendorIDSAP As String = "", Optional ByVal StaffCode As String = "", _
                                 Optional ByVal Email As String = "", Optional ByVal TelNo As String = "" _
                                 , Optional ByVal FaxNo As String = "", Optional ByVal StaffEmail As String = "" _
                                 , Optional ByVal StaffTelNo As String = "", Optional ByVal StaffMobileNo As String = "" _
                                 , Optional ByVal Address As String = "", Optional ByVal ZipCode As String = "" _
                                 , Optional ByVal ProjectType As String = "", Optional ByVal ProviceID As String = "" _
                                 , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(V.VENDOR_CODE)", VendorCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.VENDOR_NAME)", VendorName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.VENDOR_CODE_SAP)", VendorIDSAP.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.CONTACT_NAME)", ContactName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.EMAIL)", Email.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.TEL_NO)", TelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.FAX_NO)", FaxNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.ADDRESS)", ZipCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.ZIP_CODE)", Address.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "V.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "V.PROVINCE_ID", ProviceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria2, "UPPER(STAFF_NAME)", StaffName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(STAFF_CODE)", StaffCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(TEL_NO)", StaffTelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(EMAIL)", StaffEmail.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(MOBILE_NO)", StaffMobileNo.ToUpper(), DBUTIL.FieldTypes.ftText)

            SQL = "SELECT V.*, P.PROVINCE_NAME " & _
                  " FROM VENDORS V,REF_PROVINCES P WHERE" & _
                  " V.PROVINCE_ID = P.PROVINCE_ID(+)"
            If Criteria2 <> "" Then SQL &= " AND V.VENDOR_CODE IN (SELECT VENDOR_CODE FROM VENDOR_STAFFS WHERE " & Criteria2 & ")"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY V.VENDOR_CODE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchWorker(Optional ByVal VendorCode As String = "", Optional ByVal VendorName As String = "", _
                                     Optional ByVal StaffName As String = "", Optional ByVal ContactName As String = "", _
                                     Optional ByVal VendorIDSAP As String = "", Optional ByVal StaffCode As String = "", _
                                     Optional ByVal Email As String = "", Optional ByVal TelNo As String = "" _
                                     , Optional ByVal FaxNo As String = "", Optional ByVal StaffEmail As String = "" _
                                     , Optional ByVal StaffTelNo As String = "", Optional ByVal StaffMobileNo As String = "" _
                                     , Optional ByVal Address As String = "", Optional ByVal ZipCode As String = "" _
                                     , Optional ByVal ProjectType As String = "", Optional ByVal ProviceID As String = "" _
                                     , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(V.WK_CODE)", VendorCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.WORKER_NAME)", VendorName.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.VENDOR_CODE_SAP)", VendorIDSAP.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.CONTACT_NAME)", ContactName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.EMAIL)", Email.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.TEL_NO)", TelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.FAX_NO)", FaxNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.ADDRESS)", ZipCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.ZIP_CODE)", Address.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "V.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteria(Criteria, "V.PROVINCE_ID", ProviceID, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteria(Criteria2, "UPPER(STAFF_NAME)", StaffName.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(STAFF_CODE)", StaffCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(TEL_NO)", StaffTelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(EMAIL)", StaffEmail.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(MOBILE_NO)", StaffMobileNo.ToUpper(), DBUTIL.FieldTypes.ftText)

            SQL = "SELECT WK_ID, WK_CODE, WORKER_NAME, TEL_NO, MOBILE_NO, EMAIL, GROUP_ID, GROUP_NAME, PROJECT_TYPE, PROJECT_TYPE_DESC FROM  V_WORKER_DETAIL V"
            'If Criteria2 <> "" Then SQL &= " AND V.VENDOR_CODE IN (SELECT VENDOR_CODE FROM VENDOR_STAFFS WHERE " & Criteria2 & ")"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY V.WK_CODE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEventCalendar(Optional ByVal VendorCode As String = "", Optional ByVal VendorName As String = "", _
                                     Optional ByVal StaffName As String = "", Optional ByVal ContactName As String = "", _
                                     Optional ByVal VendorIDSAP As String = "", Optional ByVal StaffCode As String = "", _
                                     Optional ByVal Email As String = "", Optional ByVal TelNo As String = "" _
                                     , Optional ByVal FaxNo As String = "", Optional ByVal StaffEmail As String = "" _
                                     , Optional ByVal StaffTelNo As String = "", Optional ByVal StaffMobileNo As String = "" _
                                     , Optional ByVal Address As String = "", Optional ByVal ZipCode As String = "" _
                                     , Optional ByVal ProjectType As String = "", Optional ByVal ProviceID As String = "" _
                                     , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(NAME)", VendorCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.WORKER_NAME)", VendorName.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.VENDOR_CODE_SAP)", VendorIDSAP.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.CONTACT_NAME)", ContactName.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.EMAIL)", Email.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.TEL_NO)", TelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.FAX_NO)", FaxNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.ADDRESS)", ZipCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "UPPER(V.ZIP_CODE)", Address.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria, "V.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteria(Criteria, "V.PROVINCE_ID", ProviceID, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteria(Criteria2, "UPPER(STAFF_NAME)", StaffName.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(STAFF_CODE)", StaffCode.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(TEL_NO)", StaffTelNo.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(EMAIL)", StaffEmail.ToUpper(), DBUTIL.FieldTypes.ftText)
            'DB.AddCriteria(Criteria2, "UPPER(MOBILE_NO)", StaffMobileNo.ToUpper(), DBUTIL.FieldTypes.ftText)

            SQL = "SELECT  ID, NAME, EVENTSTART, EVENTEND, DATE_UPDATED FROM EVENT WHERE 1=1 "
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EVENTSTART"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteGroup(Optional ByVal SiteGrpID As String = "", Optional ByVal SiteGrpName As String = "", _
                                 Optional ByVal Province As String = "", Optional ByVal Remark As String = "", _
                                 Optional ByVal SiteID As String = "", Optional ByVal SiteName As String = "", _
                                 Optional ByVal ProjectType As String = "", Optional ByVal SiteType As String = "", _
                                 Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SG.SITE_GROUP_ID", SiteGrpID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SG.SITE_GROUP_NAME)", SiteGrpName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SG.PROVINCE_ID", Province, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SG.REMARK)", Remark.ToUpper(), DBUTIL.FieldTypes.ftText)

            DB.AddCriteria(Criteria2, "UPPER(SGL.SITE_ID)", SiteID.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(S.SITE_NAME)", SiteName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "S.SITE_TYPE", SiteType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria2, "S.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT SG.*, P.PROVINCE_NAME " & _
                  " FROM SITE_GROUPS SG,REF_PROVINCES P WHERE" & _
                  " SG.PROVINCE_ID = P.PROVINCE_ID(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If Criteria2 <> "" Then SQL &= " AND SG.SITE_GROUP_ID IN (SELECT SGL.SITE_GROUP_ID FROM SITE_GROUP_LISTS SGL,SITES S " & _
            "WHERE SGL.SITE_ID=S.SITE_ID(+) AND " & Criteria2 & ")"
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SG.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchProvince(Optional ByVal ProvinceID As String = "", Optional ByVal ProvinceName As String = "", _
                             Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PROVINCE_ID", ProvinceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PROVINCE_NAME", ProvinceName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_PROVINCES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PROVINCE_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchRegion(Optional ByVal RegionID As String = "", Optional ByVal RegionName As String = "", _
                             Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "REGION_ID", RegionID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "REGION_NAME", RegionName, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_REGIONS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY REGION_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchDocumentCategory(Optional ByVal DCID As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "DC_ID", DCID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM REF_DOCUMENT_CATEGORY"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "DC_ID", OrderBy)

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSoftware(Optional ByVal SoftwareID As String = "", Optional ByVal SoftwareType As String = "", _
                                   Optional ByVal SoftwareName As String = "", Optional ByVal SoftwareDetail As String = "", _
                                   Optional ByVal VersionNo As String = "", Optional ByVal DeveloperName As String = "", _
                                   Optional ByVal NetworkType As String = "", Optional ByVal LaunchDate As String = "", _
                                   Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "SOFTWARE_ID", SoftwareID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SOFTWARE_TYPE", SoftwareType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SOFTWARE_NAME", SoftwareName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SOFTWARE_DETAIL", SoftwareDetail, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "VERSION_NO", VersionNo, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "DEVELOPER_NAME", DeveloperName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "LAUNCH_DATE", AppDateValue(LaunchDate), DBUTIL.FieldTypes.ftDate)

            SQL = "SELECT * FROM SOFTWARES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "SOFTWARE_ID", OrderBy)

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Site"

#Region "Search Site"

    Public Function SearchSiteData(Optional ByVal SiteID As String = "", Optional ByVal SAPPlantCode As String = "" _
    , Optional ByVal SiteName As String = "", Optional ByVal SAPSiteName As String = "", Optional ByVal SiteType As String = "" _
    , Optional ByVal ProjectType As String = "", Optional ByVal BranchID As String = "", Optional ByVal SiteStatus As String = "" _
    , Optional ByVal PlanInstallDateF As String = "", Optional ByVal PlanInstallDateT As String = "" _
    , Optional ByVal InstallDateF As String = "", Optional ByVal InstallDateT As String = "", Optional ByVal OwnerName As String = "" _
    , Optional ByVal PlanOpenDateF As String = "", Optional ByVal PlanOpenDateT As String = "", Optional ByVal OpenDateF As String = "" _
    , Optional ByVal OpenDateT As String = "" _
    , Optional ByVal ContactName As String = "", Optional ByVal CallSiteName As String = "", Optional ByVal SalesArea As String = "" _
    , Optional ByVal TaxpayerID As String = "", Optional ByVal Latitude As String = "", Optional ByVal Longitude As String = "" _
    , Optional ByVal CostCenter As String = "", Optional ByVal CloseDateF As String = "", Optional ByVal CloseDateT As String = "" _
    , Optional ByVal Remark As String = "", Optional ByVal SiteDetail As String = "", Optional ByVal Address As String = "" _
    , Optional ByVal Province As String = "", Optional ByVal TelNo As String = "", Optional ByVal FaxNo As String = "" _
    , Optional ByVal EMail As String = "", Optional ByVal SAPSiteID As String = "", Optional ByVal SoldToSAP As String = "", Optional ByVal ShipToSAP As String = "" _
    , Optional ByVal NetworkType As String = "", Optional ByVal NetworkInstallDateF As String = "", Optional ByVal NetworkInstallDateT As String = "" _
    , Optional ByVal LineInstallDateF As String = "", Optional ByVal LineInstallDateT As String = "", Optional ByVal LinkTestDateF As String = "" _
    , Optional ByVal LinkTestDateT As String = "", Optional ByVal NetworkProtocol As String = "" _
    , Optional ByVal IPStart As String = "", Optional ByVal IPEnd As String = "", Optional ByVal AccessPoint As String = "" _
    , Optional ByVal RouterIP As String = "", Optional ByVal RouterModel As String = "", Optional ByVal DialNo As String = "" _
    , Optional ByVal DialUser As String = "", Optional ByVal DialPassword As String = "", Optional ByVal Software As String = "" _
    , Optional ByVal SWInstallDateF As String = "", Optional ByVal SWInstallDateT As String = "", Optional ByVal RemarkSW As String = "" _
    , Optional ByVal ProjectTypeHW As String = "", Optional ByVal SystemName As String = "", Optional ByVal OwnerType As String = "" _
    , Optional ByVal NetworkTypeHW As String = "", Optional ByVal HWPlanInstallDateF As String = "", Optional ByVal HWPlanInstallDateT As String = "" _
    , Optional ByVal HWPlanOpenDateF As String = "", Optional ByVal HWPlanOpenDateT As String = "" _
    , Optional ByVal HWInstallDateF As String = "", Optional ByVal HWInstallDateT As String = "" _
    , Optional ByVal HWOpenDateF As String = "", Optional ByVal HWOpenDateT As String = "", Optional ByVal SystemStatus As String = "" _
    , Optional ByVal SerialNo As String = "", Optional ByVal EquipmentName As String = "", Optional ByVal EquipmentType As String = "" _
    , Optional ByVal EquipmentStatus As String = "", Optional ByVal Location As String = "", Optional ByVal Network As String = "" _
    , Optional ByVal System As String = "", Optional ByVal EQInstallDateF As String = "", Optional ByVal EQInstallDateT As String = "" _
    , Optional ByVal ServiceNo As String = "", Optional ByVal ProjectTypeSV As String = "", Optional ByVal CallBy As String = "" _
    , Optional ByVal CallDetail As String = "", Optional ByVal ServiceType As String = "", Optional ByVal ServiceStatus As String = "" _
    , Optional ByVal SVCloseDateF As String = "", Optional ByVal SVCloseDateT As String = "", Optional ByVal MovementDateF As String = "" _
    , Optional ByVal MovementDateT As String = "", Optional ByVal MovementType As String = "", Optional ByVal SerialNoEQH As String = "" _
    , Optional ByVal PartNo As String = "", Optional ByVal EquipName As String = "", Optional ByVal SiteRefID As String = "" _
    , Optional ByVal SiteRefName As String = "", Optional ByVal ServiceNoEH As String = "", Optional ByVal ServiceTypeEH As String = "" _
    , Optional ByVal ServiceStatusEH As String = "", Optional ByVal SoftwareStatus As String = Nothing _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", Optional ByVal AdvanceSearch As Boolean = False, Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable

        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = "", Criteria3 As String = "", Criteria4 As String = "" _
        , Criteria5 As String = "", Criteria6 As String = "", Criteria7 As String = "", Criteria8 As String = "", Criteria9 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(S.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(S.SAP_PLANT_CODE)", SAPPlantCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria, "UPPER(S.SAP_SITE_ID1)", "UPPER(S.SAP_SITE_ID2)", SAPSiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(S.SAP_NAME)", SAPSiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "S.SITE_TYPE", SiteType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "S.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(S.BRANCH_ID)", BranchID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "S.STATUS", SiteStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "S.PLAN_INSTALL_DATE", AppDateValue(PlanInstallDateF), AppDateValue(PlanInstallDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "S.INSTALL_DATE", AppDateValue(InstallDateF), AppDateValue(InstallDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(S.OWNER_NAME)", OwnerName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "S.PLAN_OPEN_DATE", AppDateValue(PlanOpenDateF), AppDateValue(PlanOpenDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "S.OPEN_DATE", AppDateValue(OpenDateF), AppDateValue(OpenDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(S.CONTACT_NAME)", ContactName.ToUpper, DBUTIL.FieldTypes.ftText)

            If AdvanceSearch Then
                DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME2)", CallSiteName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "S.SALE_AREA", SalesArea, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.TAX_ID)", TaxpayerID.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.LATITUDE)", Latitude.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.LONGITUDE)", Longitude.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.COST_CENTER)", CostCenter.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteriaRange(Criteria, "S.CLOSED_DATE", AppDateValue(CloseDateF), AppDateValue(CloseDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria, "UPPER(S.REMARK)", Remark.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.SITE_DESC)", SiteDetail.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.ADDRESS)", Address.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "S.PROVINCE_ID", Province, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "UPPER(S.TEL_NO)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.FAX_NO)", FaxNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.EMAIL)", EMail.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.SAP_SOLD_TO)", SoldToSAP.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(S.SAP_SHIP_TO)", ShipToSAP.ToUpper, DBUTIL.FieldTypes.ftText)

                DB.AddCriteria(Criteria2, "SN.NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftText)
                DB.AddCriteriaRange(Criteria2, "SN.NETWORK_INSTALL_DATE", AppDateValue(NetworkInstallDateF), AppDateValue(NetworkInstallDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteriaRange(Criteria2, "SN.LINE_INSTALL_DATE", AppDateValue(LineInstallDateF), AppDateValue(LineInstallDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteriaRange(Criteria2, "SN.LINK_TEST_DATE", AppDateValue(LinkTestDateF), AppDateValue(LinkTestDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria2, "SN.NET_PROTOCOL", NetworkProtocol, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.IP_START)", IPStart.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.IP_END)", IPEnd.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.ACCESS_POINT)", AccessPoint.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.ROUTER_IP)", RouterIP.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "SN.MODEL_ID", RouterModel, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria2, "UPPER(SN.DIAL_NO)", DialNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.DIAL_USER)", DialUser.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria2, "UPPER(SN.DIAL_PASSWORD)", DialPassword.ToUpper, DBUTIL.FieldTypes.ftText)

                DB.AddCriteria(Criteria3, "SOFTWARE_ID", Software, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteriaRange(Criteria3, "INSTALL_DATE", AppDateValue(SWInstallDateF), AppDateValue(SWInstallDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria3, "UPPER(SOFTWARE_REMARK)", RemarkSW.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria3, "SYSTEM_STATUS", SoftwareStatus, DBUTIL.FieldTypes.ftNumeric)

                DB.AddCriteria(Criteria4, "PROJECT_TYPE", ProjectTypeHW, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria4, "UPPER(SYSTEM_NAME)", SystemName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria4, "OWNER_TYPE", OwnerType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria4, "NETWORK_TYPE", NetworkTypeHW, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteriaRange(Criteria4, "PLAN_INSTALL_DATE", AppDateValue(HWPlanInstallDateF), AppDateValue(HWPlanInstallDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteriaRange(Criteria4, "PLAN_OPEN_DATE", AppDateValue(HWPlanOpenDateF), AppDateValue(HWPlanOpenDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteriaRange(Criteria4, "INSTALL_DATE", AppDateValue(HWInstallDateF), AppDateValue(HWInstallDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteriaRange(Criteria4, "OPEN_DATE", AppDateValue(HWOpenDateF), AppDateValue(HWOpenDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria4, "SYSTEM_STATUS", SystemStatus, DBUTIL.FieldTypes.ftNumeric)

                DB.AddCriteria(Criteria5, "UPPER(EQ.SERIAL_NO)", SerialNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria2Condi(Criteria5, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipmentName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria5, "EQ.EQUIPMENT_TYPE", EquipmentType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria5, "EQ.EQUIPMENT_STATUS", EquipmentStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria5, "SE.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria5, "SE.NETWORK_ID", Network, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria5, "SE.SYSTEM_ID", System, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteriaRange(Criteria5, "SE.INSTALL_DATE", AppDateValue(EQInstallDateF), AppDateValue(EQInstallDateT), DBUTIL.FieldTypes.ftDate)

                DB.AddCriteria(Criteria6, "UPPER(SV.SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria6, "SV.PROJECT_TYPE", ProjectTypeSV, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria6, "UPPER(SV.INFORMER_NAME)", CallBy.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria6, "UPPER(SV.CALL_DETAIL)", CallDetail.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria6, "SV.SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria6, "SV.SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteriaRange(Criteria6, "SV.CLOSE_DATE", AppDateValue(SVCloseDateF), AppDateValue(SVCloseDateT), DBUTIL.FieldTypes.ftDate)

                DB.AddCriteriaRange(Criteria7, "EM.TRANS_DATE", AppDateValue(MovementDateF), AppDateValue(MovementDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria7, "EM.MOVEMENT_TYPE", MovementType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria7, "UPPER(EQ.SERIAL_NO)", SerialNoEQH.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria7, "UPPER(EQ.PART_NO)", PartNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria2Condi(Criteria7, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria7, "UPPER(EM.SITE_ID_OLD)", SiteRefID.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria2Condi(Criteria7, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME2)", SiteRefName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria7, "UPPER(SV.SERVICE_NO)", ServiceNoEH.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria7, "SV.SERVICE_TYPE", ServiceTypeEH, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria7, "SV.SERVICE_STATUS", ServiceStatusEH, DBUTIL.FieldTypes.ftNumeric)

                DB.AddCriteriaRange(Criteria8, "EM.TRANS_DATE", AppDateValue(MovementDateF), AppDateValue(MovementDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria8, "EM.MOVEMENT_TYPE", MovementType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria8, "UPPER(EQ.SERIAL_NO)", SerialNoEQH.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria8, "UPPER(EQ.PART_NO)", PartNo.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria2Condi(Criteria8, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria8, "UPPER(EM.SITE_ID)", SiteRefID.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria2Condi(Criteria8, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME2)", SiteRefName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria8, "UPPER(SV.SERVICE_NO)", ServiceNoEH.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria8, "SV.SERVICE_TYPE", ServiceTypeEH, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria8, "SV.SERVICE_STATUS", ServiceStatusEH, DBUTIL.FieldTypes.ftNumeric)
            End If
            '"SL.LOCATION_ID, SL.LOCATION_NAME, SL.LOCATION_DESC, " & _

            'SQL = "SELECT S.*, ST.SITE_TYPE_DESC, P.PROVINCE_NAME,P.REGION_ID, SA.SALE_AREA_NAME, " & _
            '          "SLAP.PROJECT_TYPE, SLAP.PROFILE_NAME, SLAP.ACTIVE_FLAG, " & _
            '          "AST.PROJECT_TYPE_DESC, SP.PLANT_EN_DESC, SP.PLANT_TH_DESC " & _
            '      "FROM SITES S,REF_SITE_TYPES ST,SITE_LOCATIONS SL,REF_PROJECT_TYPES AST,REF_PROVINCES P,REF_SALE_AREAS SA" & _
            '      ",SLA_PROFILES SLAP,SAP_PLANTS SP WHERE S.SITE_TYPE = ST.SITE_TYPE(+) " & _
            '          "AND S.SITE_ID = SL.SITE_ID(+) " & _
            '          "AND SL.PROJECT_TYPE = AST.PROJECT_TYPE(+) " & _
            '          "AND S.PROVINCE_ID = P.PROVINCE_ID(+) " & _
            '          "AND S.SALE_AREA = SA.SALE_AREA(+) " & _
            '          "AND S.SLA_PROFILE_ID = SLAP.SLA_PROFILE_ID(+) " & _
            '          "AND S.SAP_PLANT_CODE = SP.PLANT_CODE(+) "
            SQL = "SELECT S.*, ST.SITE_TYPE_DESC, P.PROVINCE_NAME,P.REGION_ID, SA.SALE_AREA_NAME, " & _
                      "SLAP.PROJECT_TYPE, SLAP.PROFILE_NAME, SLAP.ACTIVE_FLAG, " & _
                      "SP.PLANT_EN_DESC, SP.PLANT_TH_DESC, STS.SITE_STATUS_DESC " & _
                  "FROM SITES S,REF_SITE_TYPES ST,REF_PROVINCES P,REF_SALE_AREAS SA" & _
                  ",SLA_PROFILES SLAP,SAP_PLANTS SP, REF_SITE_STATUS STS" & _
                  " WHERE S.SITE_TYPE = ST.SITE_TYPE(+) " & _
                      "AND S.PROVINCE_ID = P.PROVINCE_ID(+) " & _
                      "AND S.SALE_AREA = SA.SALE_AREA(+) " & _
                      "AND S.SLA_PROFILE_ID = SLAP.SLA_PROFILE_ID(+) " & _
                      "AND S.STATUS = STS.SITE_STATUS(+) " & _
                      "AND S.SAP_PLANT_CODE = SP.PLANT_CODE(+) "
            If Criteria <> "" Then SQL &= " AND " & Criteria

            If AdvanceSearch Then
                If Criteria2 <> "" Then SQL &= " AND S.SITE_ID IN (SELECT SN.SITE_ID FROM SITE_NETWORKS SN WHERE " & Criteria2 & ")"
                If Criteria3 <> "" Then SQL &= " AND S.SITE_ID IN (SELECT SITE_ID FROM SITE_SOFTWARES WHERE " & Criteria3 & ")"
                If Criteria4 <> "" Then SQL &= " AND S.SITE_ID IN (SELECT SITE_ID FROM SITE_SYSTEMS WHERE " & Criteria4 & ")"
                If Criteria5 <> "" Then SQL &= " AND S.SITE_ID IN (SELECT SE.SITE_ID FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ " & _
                "WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND " & Criteria5 & ")"
                If Criteria6 <> "" Then SQL &= " AND S.SITE_ID IN (SELECT SV.SITE_ID FROM SERVICES SV WHERE " & Criteria6 & ")"
                If Criteria7 <> "" Then Criteria9 = " S.SITE_ID IN (SELECT EM.SITE_ID FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ" & _
                ",SITES S,SERVICES SV WHERE " & _
                " EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID_OLD=S.SITE_ID(+) " & _
                " AND EM.SERVICE_ID=SV.SERVICE_ID(+) AND " & Criteria7 & ")"
                If Criteria8 <> "" Then
                    If Criteria9 <> "" Then Criteria9 &= " OR "
                    Criteria9 &= " S.SITE_ID IN (SELECT EM.SITE_ID_OLD FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,SITES S,SERVICES SV" & _
                    " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID=S.SITE_ID(+) " & _
                "AND EM.SERVICE_ID=SV.SERVICE_ID(+) AND " & Criteria8 & ")"
                End If
                If Criteria9 <> "" Then SQL &= " AND (" & Criteria9 & ")"
            End If

            SQL &= " ORDER BY " & IIf(OrderBy = "", "S.DATE_UPDATED DESC", OrderBy)

            DB.OpenDT(DT, SQL, Conn, Trans)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteGroupList(Optional ByVal SiteGroupID As String = "", Optional ByVal SiteID As String = "", _
    Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable

        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria

            DB.AddCriteria(Criteria, "GL.SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "GL.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT GL.SITE_GROUP_ID,S.*, ST.SITE_TYPE_DESC, P.PROVINCE_NAME,P.REGION_ID, SA.SALE_AREA_NAME, " & _
                      "SLAP.PROJECT_TYPE, SLAP.PROFILE_NAME, SLAP.ACTIVE_FLAG, " & _
                      "AST.PROJECT_TYPE_DESC, SP.PLANT_EN_DESC, SP.PLANT_TH_DESC,SG.SITE_GROUP_NAME " & _
                  "FROM SITE_GROUP_LISTS GL,SITES S,REF_SITE_TYPES ST,REF_PROJECT_TYPES AST,REF_PROVINCES P,REF_SALE_AREAS SA" & _
                  ",SLA_PROFILES SLAP,SAP_PLANTS SP,SITE_GROUPS SG WHERE GL.SITE_ID=S.SITE_ID AND S.SITE_TYPE = ST.SITE_TYPE(+) " & _
                      "AND S.PROJECT_TYPE = AST.PROJECT_TYPE(+) " & _
                      "AND S.PROVINCE_ID = P.PROVINCE_ID(+) " & _
                      "AND S.SALE_AREA = SA.SALE_AREA(+) " & _
                      "AND S.SLA_PROFILE_ID = SLAP.SLA_PROFILE_ID(+) " & _
                      "AND S.SAP_PLANT_CODE = SP.PLANT_CODE(+) " & _
                      "AND GL.SITE_GROUP_ID = SG.SITE_GROUP_ID(+) "
            If Criteria <> "" Then SQL &= " AND " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "GL.DATE_UPDATED DESC", OrderBy)

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteNetwork(Optional ByVal SiteID As String = "", Optional ByVal NetworkID As String = "", _
    Optional ByVal NetworkType As String = "", Optional ByVal NetProtocal As String = "", Optional ByVal ModelID As String = "", _
    Optional ByVal IPStart As String = "", Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SN.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SN.NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SN.NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SN.NET_PROTOCAL", NetProtocal, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SN.MODEL_ID", ModelID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SN.IP_START)", IPStart.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = " SELECT SN.*, NT.NETWORK_TYPE_DESC, NP.NET_PROTOCOL_DESC, M.MODEL_NAME,NT.NETWORK_TYPE_DESC || ' ' || SN.IP_START AS NETWORK_NAME " & _
                ",SN.SITE_ID || ' - ' || NT.NETWORK_TYPE_DESC || ' ' || SN.IP_START AS NETWORK_NAME2" & _
                  " FROM SITE_NETWORKS SN LEFT OUTER JOIN REF_NETWORK_TYPES NT ON SN.NETWORK_TYPE = NT.NETWORK_TYPE" & _
                  " LEFT OUTER JOIN REF_NET_PROTOCOLS NP ON SN.NET_PROTOCOL = NP.NET_PROTOCOL" & _
                  " LEFT OUTER JOIN REF_MODELS M ON SN.MODEL_ID = M.MODEL_ID"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SN.SITE_ID, SN.NETWORK_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteSoftware(Optional ByVal SiteID As String = "", Optional ByVal SoftwareID As String = "", _
                                           Optional ByVal InstallDate As String = "", Optional ByVal SoftwareRemark As String = "", _
                                           Optional ByVal SystemStatus As String = Nothing, _
                                           Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "SS.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SS.SOFTWARE_ID", SoftwareID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SS.INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "SS.SYSTEM_STATUS", SoftwareID, DBUTIL.FieldTypes.ftNumeric)

            SQL = " SELECT SS.*, S.*, ST.SOFTWARE_TYPE_DESC, RSS.SYSTEM_STATUS_DESC " & _
                  " FROM SITE_SOFTWARES SS, SOFTWARES S, " & _
                  " REF_SOFTWARE_TYPES ST, REF_SYSTEM_STATUS RSS" & _
                  " WHERE SS.SOFTWARE_ID = S.SOFTWARE_ID(+) AND" & _
                  " S.SOFTWARE_TYPE = ST.SOFTWARE_TYPE(+) AND SS.SYSTEM_STATUS=RSS.SYSTEM_STATUS(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria

            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SS.SITE_ID "
            End If

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteSystem(Optional ByVal SiteID As String = "", Optional ByVal SystemID As String = "", _
                                         Optional ByVal ProjectType As String = "", Optional ByVal SystemName As String = "", _
                                         Optional ByVal OwnerType As String = "", Optional ByVal NetworkType As String = "", _
                                         Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SS.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SS.SYSTEM_ID", SystemID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SS.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SS.SYSTEM_NAME", SystemName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SS.OWNER_TYPE", OwnerType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SS.NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftNumeric)

            SQL = " SELECT SS.*, PT.PROJECT_TYPE_DESC, NT.NETWORK_TYPE_DESC, WT.OWNER_TYPE_DESC,RSS.SYSTEM_STATUS_DESC" & _
                    ",SS.SITE_ID || ' - ' || SS.SYSTEM_NAME AS SYSTEM_NAME2 " & _
                  " FROM SITE_SYSTEMS SS, REF_PROJECT_TYPES PT," & _
                  " REF_NETWORK_TYPES NT, REF_OWNER_TYPES WT,REF_SYSTEM_STATUS RSS" & _
                  " WHERE SS.PROJECT_TYPE = PT.PROJECT_TYPE(+) AND" & _
                  " SS.NETWORK_TYPE = NT.NETWORK_TYPE(+) AND" & _
                  " SS.OWNER_TYPE = WT.OWNER_TYPE(+) AND SS.SYSTEM_STATUS=RSS.SYSTEM_STATUS(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria

            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SS.SITE_ID, SS.SYSTEM_ID "
            End If

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteEquipment(Optional ByVal SiteID As String = "", Optional ByVal EquipmentID As String = "", _
                                        Optional ByVal LocationaID As String = "", Optional ByVal NetworkID As String = "", _
                                        Optional ByVal InstallDate As String = "", Optional ByVal SystemID As String = "", _
                                        Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SE.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SE.EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.LOCATION_ID", LocationaID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.SYSTEM_ID", SystemID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)

            'SQL = " SELECT SE.*,SS.*, E.*, SL.*,SN.*" & _
            '      " FROM SITE_EQUIPMENTS SE, SITE_SYSTEMS SS, EQUIPMENTS E, SITE_LOCATIONS SL, SITE_NETWORKS SN" & _
            '      " WHERE SE.SITE_ID = SS.SITE_ID(+) AND SE.SYSTEM_ID = SS.SYSTEM_ID(+) AND" & _
            '      " SE.EQUIPMENT_ID = E.EQUIPMENT_ID(+) AND" & _
            '      " SE.LOCATION_ID = SL.LOCATION_ID(+) AND" & _
            '      " SE.NETWORK_ID = SN.NETWORK_ID(+)"
            SQL = "SELECT SE.SITE_ID, SE.EQUIPMENT_ID, SE.LOCATION_ID, SE.NETWORK_ID, SE.INSTALL_DATE, SE.DATE_UPDATED, SE.USER_UPDATED, SE.SYSTEM_ID, " & _
                  "SS.PROJECT_TYPE, SS.SYSTEM_NAME, SS.OWNER_TYPE, SS.NETWORK_TYPE, " & _
                  "E.SERIAL_NO, E.BARCODE_NO, E.PART_NO, E.SAP_MAT_CODE, E.BRAND_ID, E.MODEL_ID, " & _
                  "E.SHORT_DESC, E.EQUIPMENT_DESC, E.EQUIPMENT_SPEC, E.UOM, E.QUANTITY, E.UNIT_COST, E.TOTAL_COST, E.EQUIPMENT_STATUS, " & _
                  "E.EQUIPMENT_TYPE, E.UNIT_ID, E.WA_DATE_START, E.WA_DATE_END, E.WARRANTY_TYPE, E.EQUIP_SET_FLAG, E.PM_DATE,E.VENDOR_RESPONSE, " & _
                  "SL.LOCATION_NAME, SL.LOCATION_DESC, " & _
                  "SN.NET_PROTOCOL, SN.IP_START, SN.IP_END, SN.ACCESS_POINT, SN.ROUTER_IP, SN.DIAL_NO, SN.DIAL_USER, SN.DIAL_PASSWORD" & _
                  ", SN.NETWORK_INSTALL_DATE, SN.LINE_INSTALL_DATE, SN.LINK_TEST_DATE,V.VENDOR_NAME,V2.VENDOR_NAME AS VENDOR_RESPONSE_NAME " & _
                  "FROM SITE_EQUIPMENTS SE, SITE_SYSTEMS SS, EQUIPMENTS E, SITE_LOCATIONS SL, SITE_NETWORKS SN,VENDORS V,VENDORS V2 " & _
                  "WHERE SE.SITE_ID = SS.SITE_ID(+) AND SE.SYSTEM_ID = SS.SYSTEM_ID(+) AND " & _
                    "SE.EQUIPMENT_ID = E.EQUIPMENT_ID(+) AND " & _
                    "SE.SITE_ID=SL.SITE_ID(+) AND SE.LOCATION_ID = SL.LOCATION_ID(+) AND " & _
                    "SE.SITE_ID=SN.SITE_ID(+) AND SE.NETWORK_ID = SN.NETWORK_ID(+) AND E.VENDOR_CODE=V.VENDOR_CODE(+)" & _
                    "AND E.VENDOR_RESPONSE=V2.VENDOR_CODE(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "SE.SITE_ID, SE.EQUIPMENT_ID", OrderBy)

            DB.OpenDT(DT, SQL, Conn, Trans)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngSiteEquipment(ByVal op As Integer, ByVal SiteID As String, ByVal EquipmentID As String _
    , Optional ByVal LocationID As String = Nothing, Optional ByVal NetworkID As String = Nothing _
    , Optional ByVal SystemID As String = Nothing, Optional ByVal InstallDate As String = Nothing, _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_ID", SystemID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_EQUIPMENTS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSLAProfile(Optional ByVal SLAProfileID As String = "", Optional ByVal SLAProfileName As String = "", _
     Optional ByVal ProjectType As String = "", Optional ByVal VendorName As String = "", Optional ByVal ActiveFlag As String = "" _
     , Optional ByVal VendorCode As String = "", Optional ByVal SLAType As String = "", Optional ByVal SeverityLevel As String = "" _
     , Optional ByVal ResolutionTimeF As String = "", Optional ByVal ResolutionTimeT As String = "" _
     , Optional ByVal ResponseTimeF As String = "", Optional ByVal ResponseTimeT As String = "" _
     , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SLA.SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SLA.SLA_TYPE", SLAType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SLA.PROFILE_NAME)", SLAProfileName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.VENDOR_NAME)", VendorName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(V.VENDOR_CODE)", VendorCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SLA.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SLA.ACTIVE_FLAG", ActiveFlag, DBUTIL.FieldTypes.ftText)

            DB.AddCriteria(Criteria2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria2, "(RESPONSE_TIME/1440)", ResponseTimeF, ResponseTimeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria2, "(RESOLUTION_TIME/1440)", ResolutionTimeF, ResolutionTimeT, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT SLA.*, PT.PROJECT_TYPE_DESC,V.VENDOR_NAME,DECODE(SLA.ACTIVE_FLAG,'Y','Enable','Disable') ACTIVE_STATUS_DESC,DECODE(SLA.ACTIVE_FLAG,'Y','ใช้งาน','ยกเลิกการใช้งาน') ACTIVE_STATUS_DESC_T " & _
                    ",DECODE(SLA.SLA_TYPE,1,'Issue','Vendor') SLA_TYPE_DESC,PT.PROJECT_TYPE_DESC || ' - ' || SLA.PROFILE_NAME AS PROFILE_NAME2  " & _
                  " FROM SLA_PROFILES SLA,REF_PROJECT_TYPES PT,VENDORS V WHERE " & _
                  " SLA.PROJECT_TYPE = PT.PROJECT_TYPE(+) AND SLA.VENDOR_CODE=V.VENDOR_CODE(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If Criteria2 <> "" Then SQL &= " AND SLA.SLA_PROFILE_ID IN (SELECT SLA_PROFILE_ID FROM SLA_DETAILS WHERE " & Criteria2 & ")"
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SLA.DATE_UPDATED DESC"
            End If

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteLocation(Optional ByVal SiteID As String = "", Optional ByVal LocationID As String = "", Optional ByVal LocationName As String = "", Optional ByVal ProjectType As String = "", _
                                           Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SL.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SL.LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SL.LOCATION_NAME", LocationName, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SL.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT SL.*,SL.SITE_ID || ' - ' || SL.LOCATION_NAME AS LOCATION_NAME2 FROM SITE_LOCATIONS SL"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SL.SITE_ID, SL.LOCATION_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteEquipHistory(ByVal SiteID As String, Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "EM.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)

            'SQL = "SELECT EM.*, E.PART_NO, E.EQUIPMENT_DESC, SL.LOCATION_NAME,MT.MOVEMENT_TYPE_DESC" & _
            '",S.SITE_NAME AS SITE_NAME_REF,SV.SERVICE_NO " & _
            '      "FROM EQUIPMENT_MOVEMENTS EM, SITE_LOCATIONS SL,SITE_EQUIPMENTS SE, EQUIPMENTS E,REF_MOVEMENT_TYPES MT" & _
            '      ",SITES S,SERVICES SV " & _
            '      "WHERE EM.SITE_ID = SE.SITE_ID(+) AND EM.EQUIPMENT_ID = E.EQUIPMENT_ID(+) AND " & _
            '        " EM.EQUIPMENT_ID=SE.EQUIPMENT_ID(+) AND SE.LOCATION_ID = SL.LOCATION_ID(+) " & _
            '        " AND EM.MOVEMENT_TYPE=MT.MOVEMENT_TYPE(+) AND EM.REF_NO1=S.SITE_ID(+) AND EM.SERVICE_ID=SV.SERVICE_ID(+)"

            SQL = "SELECT EM.*,MT.MOVEMENT_TYPE_DESC,EQ.PART_NO,EQ.EQUIPMENT_DESC,S.SITE_NAME,S2.SITE_NAME AS SITE_NAME_REF,SV.SERVICE_NO" & _
                  " FROM EQUIPMENT_MOVEMENTS EM,REF_MOVEMENT_TYPES MT,EQUIPMENTS EQ,SITES S,SITES S2,SERVICES SV" & _
                  " WHERE EM.MOVEMENT_TYPE=MT.MOVEMENT_TYPE(+) AND EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+)" & _
                  " AND EM.SITE_ID=S.SITE_ID(+) AND EM.SITE_ID_OLD=S2.SITE_ID(+) AND EM.SERVICE_ID=SV.SERVICE_ID(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy <> "", OrderBy, "EM.TRANS_DATE DESC")

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSitePicture(Optional ByVal SiteID As String = "", Optional ByVal PicID As String = "", Optional ByVal PicDesc As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PIC_ID", PicID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PIC_DESC", PicDesc, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM SITE_PICTURES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "SITE_ID, PIC_ID", OrderBy)

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngSiteNetwork(ByVal op As Integer, ByVal SiteID As String, _
                                       ByRef NetworkID As String, Optional ByVal NetworkType As String = Nothing, _
                                       Optional ByVal NetProtocol As String = Nothing, Optional ByVal IPStart As String = Nothing, Optional ByVal IPEnd As String = Nothing, _
                                       Optional ByVal AccessPoint As String = Nothing, Optional ByVal RouterIP As String = Nothing, Optional ByVal ModelID As String = Nothing, _
                                       Optional ByVal DailNo As String = Nothing, Optional ByVal DailUser As String = Nothing, Optional ByVal DailPassword As String = Nothing, _
                                       Optional ByVal NetworkInstallDate As String = Nothing, Optional ByVal LineInstallDate As String = Nothing, Optional ByVal LinkTestDate As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT

                    'NetworkID = GenerateID("SITE_NETWORKS", "NETWORK_ID", usrCriteria:="SITE_ID = '" & SiteID & "'") & 
                    NetworkID = GenerateID("SITE_NETWORKS", "NETWORK_ID") & ""
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL2(op, SQL1, SQL2, "NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "NETWORK_TYPE", NetworkType, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "NET_PROTOCOL", NetProtocol, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "IP_START", IPStart, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "IP_END", IPEnd, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ACCESS_POINT", AccessPoint, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ROUTER_IP", RouterIP, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "MODEL_ID", ModelID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "DIAL_NO", DailNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DIAL_USER", DailUser, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DIAL_PASSWORD", DailPassword, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "NETWORK_INSTALL_DATE", AppDateValue(NetworkInstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "LINE_INSTALL_DATE", AppDateValue(LineInstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "LINK_TEST_DATE", AppDateValue(LinkTestDate), DBUTIL.FieldTypes.ftDate)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_NETWORKS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then NetworkID = ""
            Throw ex
        End Try
    End Function
#End Region


#Region "Manage Site"

    Public Function MngSiteData(ByVal op As Integer, ByVal SiteIDOld As String, ByVal SiteID As String, _
    Optional ByVal SiteName As String = Nothing, Optional ByVal SiteName2 As String = Nothing, Optional ByVal SiteDesc As String = Nothing, _
    Optional ByVal SiteType As String = Nothing, Optional ByVal Address As String = Nothing, Optional ByVal ProvinceID As String = Nothing, Optional ByVal SaleArea As String = Nothing, _
    Optional ByVal TaxID As String = Nothing, Optional ByVal OwnerName As String = Nothing, Optional ByVal ContactName As String = Nothing, _
    Optional ByVal TelNo As String = Nothing, Optional ByVal FaxNo As String = Nothing, Optional ByVal Email As String = Nothing, _
    Optional ByVal POSSystemID As String = Nothing, Optional ByVal BackOfficeID As String = Nothing, Optional ByVal Status As String = Nothing, _
    Optional ByVal PlanInstallDate As String = Nothing, Optional ByVal PlanOpenDate As String = Nothing, Optional ByVal InstallDate As String = Nothing, Optional ByVal OpenDate As String = Nothing, _
    Optional ByVal SLAProfileID As String = Nothing, Optional ByVal SAPPlantCode As String = Nothing, Optional ByVal SAPName As String = Nothing, Optional ByVal SAPSoldTo As String = Nothing, Optional ByVal SAPShipTo As String = Nothing, _
    Optional ByVal Latitude As String = Nothing, Optional ByVal Longitude As String = Nothing, Optional ByVal BranchCode As String = Nothing _
    , Optional ByVal ProjectType As String = Nothing, Optional ByVal Remark As String = Nothing, Optional ByVal CostCenter As String = Nothing _
    , Optional ByVal SiteGrpID As String = Nothing, Optional ByVal CloseDate As String = Nothing, Optional ByVal SAPSiteID1 As String = Nothing _
    , Optional ByVal SAPSiteID2 As String = Nothing, Optional ByVal GroupView As String = Nothing, Optional ByVal BranchID As String = Nothing, Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteIDOld, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    'SiteID = GenerateID("SITES", "SITE_ID") & ""
                End If

                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_NAME", SiteName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_NAME2", SiteName2, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_DESC", SiteDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_TYPE", SiteType, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ADDRESS", Address, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROVINCE_ID", ProvinceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SALE_AREA", SaleArea, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TAX_ID", TaxID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "OWNER_NAME", OwnerName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "CONTACT_NAME", ContactName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "FAX_NO", FaxNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EMAIL", Email, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "POS_SYSTEM_ID", POSSystemID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "BACK_OFFICE_ID", BackOfficeID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "STATUS", Status, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "PLAN_INSTALL_DATE", AppDateValue(PlanInstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "PLAN_OPEN_DATE", AppDateValue(PlanOpenDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "OPEN_DATE", AppDateValue(OpenDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_NAME", SAPName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_SOLD_TO", SAPSoldTo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_SHIP_TO", SAPShipTo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LATITUDE", Latitude, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LONGITUDE", Longitude, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BRAND_CODE", BranchCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "REMARK", Remark, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "COST_CENTER", CostCenter, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_GROUP_ID", SiteGrpID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "CLOSED_DATE", AppDateValue(CloseDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_SITE_ID1", SAPSiteID1, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_SITE_ID2", SAPSiteID2, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "GROUP_VIEWS", GroupView, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BRANCH_ID", BranchID, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITES", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then SiteID = ""
            Throw ex
        End Try
    End Function

    Public Function MngSiteLocation(ByVal op As Integer, ByVal SiteID As String, ByRef LocationID As String, Optional ByVal LocationName As String = Nothing, Optional ByVal LocationDesc As String = Nothing, Optional ByVal ProjectTypeID As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT

                    LocationID = GenerateID("SITE_LOCATIONS", "LOCATION_ID") & ""
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "LOCATION_NAME", LocationName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LOCATION_DESC", LocationDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectTypeID, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_LOCATIONS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then LocationID = ""
            Throw ex
        End Try
    End Function

    Public Function MngSitePicture(ByVal op As Integer, ByVal SiteID As String, ByRef PicID As String, Optional ByVal PicDesc As String = Nothing, Optional ByVal FileName As String = Nothing) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "PIC_ID", PicID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT

                    'PicID = GenerateID("SITE_PICTURES", "PIC_ID", usrCriteria:="SITE_ID = '" & SiteID & "'") & ""
                    PicID = GenerateID("SITE_PICTURES", "PIC_ID") & ""
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL2(op, SQL1, SQL2, "PIC_ID", PicID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "PIC_DESC", PicDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "FILE_NAME", FileName, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_PICTURES", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then PicID = ""
            Throw ex
        End Try
    End Function

    Public Function MngSiteGroupList(ByVal op As Integer, ByVal SiteGroupID As String, ByVal SiteID As String) As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                End If
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_GROUP_LISTS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function MngSiteSoftware(ByVal op As Integer, ByVal SiteID As String, ByVal SoftwareID As String, _
                                    Optional ByVal InstallDate As String = Nothing, Optional ByVal SoftwareRemark As String = Nothing, _
                                    Optional ByVal SystemStatus As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "SOFTWARE_ID", SoftwareID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL2(op, SQL1, SQL2, "SOFTWARE_ID", SoftwareID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SOFTWARE_REMARK", SoftwareRemark, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_STATUS", SystemStatus, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_SOFTWARES", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then SiteID = ""
            Throw ex
        End Try
    End Function

    Public Function MngSiteSystem(ByVal op As Integer, ByVal SiteID As String, ByVal SystemID As String, _
                                    Optional ByVal ProjectType As String = Nothing, Optional ByVal SystemName As String = Nothing, _
                                    Optional ByVal OwnerType As String = Nothing, Optional ByVal Networktype As String = Nothing, _
                                    Optional ByVal PlanInstallDate As String = Nothing, Optional ByVal PlanOpenDate As String = Nothing, _
                                    Optional ByVal InstallDate As String = Nothing, Optional ByVal OpenDate As String = Nothing, _
                                    Optional ByVal POS As String = Nothing, Optional ByVal BO As String = Nothing, _
                                    Optional ByVal SystemStatus As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "SYSTEM_ID", SystemID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                    'SystemID = GenerateID("SITE_SYSTEMS", "SYSTEM_ID", usrCriteria:="SITE_ID = '" & SiteID & "'") & ""
                    SystemID = GenerateID("SITE_SYSTEMS", "SYSTEM_ID") & ""
                    DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_ID", SystemID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_NAME", SystemName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "OWNER_TYPE", OwnerType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "NETWORK_TYPE", Networktype, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "PLAN_INSTALL_DATE", AppDateValue(PlanInstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "PLAN_OPEN_DATE", AppDateValue(PlanOpenDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "OPEN_DATE", AppDateValue(OpenDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "POS", POS, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BO", BO, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_STATUS", SystemStatus, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_SYSTEMS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then SiteID = ""
            Throw ex
        End Try
    End Function

#End Region

    'ไม่ได้ใช้ SITE_EQUIPMENT แล้วใช้ EQUIPMENTS แทน
    'Public Function MngSiteEquipment(ByVal op As Integer, ByVal SiteID As String, _
    '                               ByVal EqipmentID As String, Optional ByVal LocationID As String = Nothing, _
    '                               Optional ByVal NetworkID As String = Nothing, Optional ByVal InstallDate As String = Nothing, ) As String
    '    Dim SQL1, SQL2, SQL As String
    '    Dim Criteria As String = ""

    '    Try
    '        SQL = "" : SQL1 = "" : SQL2 = ""
    '        If op <> DBUTIL.opINSERT Then
    '            DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
    '            DB.AddCriteria(Criteria, "EQUIPMENT_ID", EqipmentID, DBUTIL.FieldTypes.ftNumeric)
    '        End If
    '        If op <> DBUTIL.opDELETE Then
    '            If op = DBUTIL.opINSERT Then
    '                op = DBUTIL.opINSERT

    '                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
    '                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_ID", EqipmentID, DBUTIL.FieldTypes.ftText)
    '            End If

    '            DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
    '        End If

    '        SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_EQUIPMENTS", Criteria, True)
    '        DB.ExecSQL(SQL)
    '        Return ""
    '    Catch ex As Exception
    '        If op = opINSERT Then NetworkID = ""
    '        Throw ex
    '    End Try
    'End Function
#End Region

#Region "Service Document"
    Public Function SearchServiceDocument(Optional ByVal SDID As String = "", Optional ByVal Desc As String = "" _
    , Optional ByVal Keyword As String = "", Optional ByVal CategoryID As String = "", Optional ByVal ProjectTypeID As String = "" _
    , Optional ByVal SiteID As String = "", Optional ByVal SiteName As String = "", Optional ByVal SUserView As String = "" _
    , Optional ByVal SGroupView As String = "", Optional ByVal UserCreate As String = "", Optional ByVal CreateDateF As String = "" _
    , Optional ByVal CreateDateT As String = "", Optional ByVal UserUpdate As String = "", Optional ByVal UpdateDateF As String = "" _
    , Optional ByVal UpdateDateT As String = "", Optional ByVal DocumentView As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SD.SD_ID", SDID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SD.SD_DESC)", Desc.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SD.KEYWORD)", Keyword.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SD.CATEGORY_ID", CategoryID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SD.PROJECT_TYPE_ID", ProjectTypeID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SD.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SD.DOC_VIEW_USER)", SUserView.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SD.DOC_VIEW_USER)", SGroupView.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SD.USER_CREATED)", UserCreate.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "UPPER(SD.DATE_CREATED)", AppDateValue(CreateDateF), AppDateValue(CreateDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(SD.USER_UPDATED)", UserUpdate.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "UPPER(SD.DATE_UPDATED)", AppDateValue(UpdateDateF), AppDateValue(UpdateDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "SD.DOC_VIEW_TYPE", DocumentView, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT SD.*, PT.PROJECT_TYPE_DESC, DC.DC_DESC, S.SITE_NAME " & _
                          "FROM SERVICE_DOCUMENTS SD INNER JOIN " & _
                          "REF_DOCUMENT_CATEGORY DC ON SD.CATEGORY_ID = DC.DC_ID LEFT OUTER JOIN " & _
                          "REF_PROJECT_TYPES PT ON SD.PROJECT_TYPE_ID = PT.PROJECT_TYPE LEFT OUTER JOIN " & _
                          "SITES S ON SD.SITE_ID = S.SITE_ID"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY " & IIf(OrderBy = "", "SD.DATE_UPDATED DESC", OrderBy)

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub MngServiceDocument(ByVal op As Integer, ByRef SDID As String, Optional ByVal Desc As String = Nothing _
    , Optional ByVal Keyword As String = Nothing, Optional ByVal FileName As String = Nothing _
    , Optional ByVal CategoryID As String = Nothing, Optional ByVal ProjectTypeID As String = Nothing _
    , Optional ByVal SiteID As String = Nothing, Optional ByVal DocViewType As String = Nothing _
    , Optional ByVal DocViewUser As String = Nothing)
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "SD_ID", SDID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    'Gen New ID
                    SDID = GenerateID("SERVICE_DOCUMENTS", "SD_ID")
                    DB.AddSQL(op, SQL1, SQL2, "SD_ID", SDID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                    DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
                Else
                    op = DBUTIL.opUPDATE
                End If

                DB.AddSQL2(op, SQL1, SQL2, "SD_DESC", Desc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "KEYWORD", Keyword, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "FILE_NAME", FileName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "CATEGORY_ID", CategoryID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE_ID", ProjectTypeID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DOC_VIEW_TYPE", DocViewType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "DOC_VIEW_USER", DocViewUser, DBUTIL.FieldTypes.ftText)

                'DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                'DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", User, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICE_DOCUMENTS", Criteria, True)
                DB.ExecSQL(SQL)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Logging"
    Public Function SqlLog(ByVal sql_text As String) As String
        sql_text = sql_text.Replace("'", "")
        'sql_text = IIf(Len(sql_text) > 4000, sql_text.Substring(0, 4000), sql_text)
        Dim sql As String = String.Format("INSERT INTO SQL_LOG (SQLTEXT) VALUES  ('{0}')", sql_text)
        Try
            DB.ExecSQL(sql)
            Return ""
        Catch ex As Exception
            Throw New DALException(ex.Message)
        End Try
    End Function
#End Region
End Class
