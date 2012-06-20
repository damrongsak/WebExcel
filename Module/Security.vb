#Region ".NET Framework Class Import"
Imports System.Security
Imports System.Security.Principal
Imports System.Threading.Thread
Imports System.Net.Mail
Imports System.Data

#End Region

Public Module Security

    Public Const RET_UNKNOWN_ERR As Integer = 0
    'Updated By Aoy 11/05/2552
    Public Const TaskMngLookup As Integer = 1
    Public Const TaskMngSoftwarePOS As Integer = 2
    Public Const TaskMngVendor As Integer = 3
    Public Const TaskMngSLAProfile As Integer = 4
    Public Const TaskMngProblemResolve As Integer = 5
    Public Const TaskMngSite As Integer = 6
    Public Const TaskMngSiteNetwork As Integer = 7
    Public Const TaskMngSiteSystem As Integer = 8
    Public Const TaskMngSiteEquipment As Integer = 9
    Public Const TaskSiteServiceHist As Integer = 10
    Public Const TaskSiteEquipHist As Integer = 11
    Public Const TaskMngSurvey As Integer = 12
    Public Const TaskMngServiceDoc As Integer = 13
    Public Const TaskMngEquipment As Integer = 14
    Public Const TaskMngStockMovement As Integer = 15
    Public Const TaskStockBalance As Integer = 16
    Public Const TaskCheckPhysicalStock As Integer = 17
    Public Const TaskMngService As Integer = 18
    Public Const TaskServiceReport As Integer = 19
    Public Const TaskImportData As Integer = 20
    Public Const TaskReconcile As Integer = 21
    Public Const TaskReport As Integer = 22
    Public Const TaskMngUser As Integer = 23
    Public Const TaskMngRole As Integer = 24
    Public Const TaskAuditLog As Integer = 25
    Public Const TaskExceptionLog As Integer = 26
    Public Const TaskMngSiteGroup As Integer = 27
    Public Const TaskDashboard As Integer = 28
    Public Const TaskMngGroup As Integer = 29

    'Public Const actRead As Integer = 1    ' ค้นหา, ดูข้อมูล
    'Public Const actModify As Integer = 2   'แก้ไขข้อมูล
    'Public Const actConfirm As Integer = 4   'ยืนยันข้อมูล

    Public Const actView As Integer = 1        ' ค้นหา, ดูข้อมูล
    Public Const actModify As Integer = 2         ' เพิ่ม แก้ไขข้อมูล
    Public Const actDelete As Integer = 4        ' ลบข้อมูล

    Public gUSER_NAME As String

    Public Const CharList As String = " 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_@#$%^&*()-+/=.,"
    Public Const CharCount As Integer = 79
    Public Const MaxKeyLen As Integer = 20

    'Public Const gBannedDuration As Integer = 10    ' minutes


    'Updated By Aoy 12/05/2552
    'Public Function CheckRole(ByVal TASK_ID As Integer, Optional ByVal Action As Integer = actView, Optional ByVal isGoNoRight As Boolean = True) As Integer
    '    Dim CurrentPriviledge As Integer = 0

    '    If gReadConfig & "" <> "Y" Then ReadConfigurations()
    '    If HttpContext.Current.Session("USER_NAME") & "" = "" Then
    '        HttpContext.Current.Response.Write("<script language=""javascript"">" & _
    '        "alert('Unable to access this page ! Please contact system administrator if you require to access.');" & _
    '        "if (opener){window.opener.location.href='../Logout.aspx'; this.close();}else if(parent)" & _
    '        "{window.parent.location.href='../Logout.aspx';} else" & _
    '        "{window.location.href='../Logout.aspx';}</script>")
    '    Else
    '        If Not IsAuthorized(TASK_ID, Action) AndAlso isGoNoRight Then
    '            'HttpContext.Current.Response.Write("<script language=""javascript"">if (parent != null" & _
    '            '" || parent != undefined){parent.location.href='../NoRight.aspx';} else if (opener != null" & _
    '            '" || opener != undefined){opener.location.href='../NoRight.aspx'; this.close();} else " & _
    '            '"{window.location.href='../NoRight.aspx';}</script>")
    '            HttpContext.Current.Response.Write("<script language=""javascript"">alert('Unable to" & _
    '                " access this page ! Please contact system administrator if you require to access.');" & _
    '                " if(opener){ window.close();}else if(parent){history.back(1);}else{history.back(1);}" & _
    '                " </script>")
    '        Else
    '            If IsAuthorized(TASK_ID, Action) Then
    '                CurrentPriviledge = 1
    '            End If
    '        End If
    '    End If
    '    Return CurrentPriviledge
    'End Function

    Public Function CheckRole(ByVal TASK_ID As Integer, Optional ByVal isGoNoRight As Boolean = True _
    , Optional ByVal act As Integer = actView, Optional ByVal chkCanViewHist As Boolean = False) As Integer
        Dim CurrentPriviledge As Integer = 0
        Dim UID As String = Trim(HttpContext.Current.Session("UID") & "")
        If gReadConfig & "" <> "Y" Then ReadConfigurations()

        If isGoNoRight AndAlso (Not IsAuthorized(TASK_ID, act) Or UID = "" Or UID <> HttpContext.Current.Request("UID") & "") Then
            'If Not (IsAuthenticated() AndAlso Val(HttpContext.Current.Session("ROLES") & "") >= PrivRole) Then

            'HttpContext.Current.Response.Redirect("../noRight.aspx")
            'HttpContext.Current.Response.Write("<script language='javascript'>window.location.href='../NoRight.aspx';</script>")
            'HttpContext.Current.Response.Write("<script language='javascript'>alert('No Right');</script>")
            'If Not chkCanViewHist OrElse (chkCanViewHist AndAlso HttpContext.Current.Request("hdnCanViewHist") & "" <> "Y") Then
            '    If HttpContext.Current.Session("UID") & "" = "" Then
            '        HttpContext.Current.Session("UID") = "NoRight"
            '    End If
            '    HttpContext.Current.Response.Write("<script language=""javascript"">if (parent != null || parent != undefined){parent.location.href='../NoRight.htm';} else if (opener != null || opener != undefined){opener.location.href='../NoRight.htm'; this.close();} else {window.location.href='../NoRight.htm';}</script>")
            '    'HttpContext.Current.Response.Write("<script language=""javascript"">if (parent != null || parent != undefined){parent.location.href='../Logout.aspx';} else if (opener != null || opener != undefined){opener.location.href='../Logout.aspx'; this.close();} else {window.location.href='../Logout.aspx';}</script>")
            'End If
            If UID <> HttpContext.Current.Request("UID") & "" Then
                If UID <> "" Then
                    BLL.InsertAudit(catBannedRequestLog, "Too many request failure!", HttpContext.Current.Session("USER_NAME") & "")
                End If
                ClearSession()
            End If
            HttpContext.Current.Response.Write("<script language=""javascript"">if (parent != null || parent != undefined)" & _
            "{parent.location.href='../NoRight.htm';} else if (opener != null || opener != undefined)" & _
            "{opener.location.href='../NoRight.htm'; this.close();} else {window.location.href='../NoRight.htm';}</script>")
        Else
        If IsAuthorized(TASK_ID, actView) Then
            CurrentPriviledge = 1
        End If
        End If

        Return CurrentPriviledge
    End Function


    Public Function CanDo(ByVal TASK_ID As Integer, ByVal Action As Integer, Optional ByVal Permits As Object = Nothing) As Boolean
        CanDo = IsAuthorized(TASK_ID, Action, Permits)
    End Function

    Public Function CannotDo(ByVal TASK_ID As Integer, ByVal Action As Integer, Optional ByVal Permits As Object = Nothing) As Boolean
        CannotDo = True
        If IsAuthorized(TASK_ID, Action, Permits) Then
            CannotDo = False
        End If
    End Function

    Public Function IsAuthorized(ByVal TASK_ID As Integer, ByVal Action As Integer, Optional ByVal Permissions As Object = Nothing) As Boolean
        Dim P, RIGHTS As Object
        P = 0
        If Not Permissions = Nothing Then
            RIGHTS = Permissions
        Else
            RIGHTS = HttpContext.Current.Session("RIGHTS") & ""
        End If
        P = Asc(Mid(RIGHTS, TASK_ID, 1) & "@") - 64
        IsAuthorized = ((Action And P) <> 0)
    End Function

    '=================================================================
    'function แสดงค่าตามตัวแปร show เช่น link file หรือ msgbox
    '   ในกรณีที่มีสิทธิ์  แต่ในกรณีไม่มีสิทธิ์จะแสดงข้อความเตือน
    Public Function ShowCanDo(ByVal TASK_ID As Integer, ByVal Action As Integer, ByVal Show As String, ByVal Permit As Object) As Object
        If IsAuthorized(TASK_ID, Action, Permit) Then
            ShowCanDo = " '" & Show & "' "
        Else
            ShowCanDo = " 'VBScript:Alert(""  ไม่มีสิทธิ์ทำงานนี้!  "")'"
        End If
    End Function

    '=================================================================
    ' Function สำหรับ เข้ารหัส รหัสผ่าน ก่อนบันทึกลงฐานข้อมูล
    Function Encrypted(ByVal Key1 As Object, ByVal Key2 As Object) As String
        Dim I As Integer
        Dim X As Integer
        Dim S As String = ""
        Key1 = Trim(CStr(Key1))
        Key2 = Trim(CStr(Key2))
        X = 55
        For I = 1 To 10
            If I > 10 - Len(Key1) Then
                X = (X + I) Xor Asc(Mid(Key1, 10 - I + 1, 1))
            Else
                X = X Xor I
            End If
            If I <= Len(Key2) Then
                X = X Xor Asc(Mid(Key2, I, 1))
            Else
                X = X Xor (I * 3)
            End If
            X = X And 127
            If X = 124 Then
                X = 125
            ElseIf (X < 32) Then
                X = X + 32
            End If
            If X = 39 Then X = 40
            S = S & Chr(X)
        Next
        Encrypted = S
    End Function

    Public Function Key2Char(ByVal num As Object) As String
        Key2Char = Mid(CharList, (num Mod CharCount) + 1, 1)
    End Function

    Public Function Key2Num(ByVal Key As String) As Integer
        Key2Num = InStr(CharList, Key) - 1
    End Function

    Public Function DecodeKey(ByVal SecretKey As String, ByVal EncodedKey As String) As String
        Dim X As Integer
        Dim S As String = "", t As String = ""
        Dim num As Integer
        Dim Data As String
        Dim I As Integer
        Try
            If Len(EncodedKey) <> MaxKeyLen + 3 Then
                DecodeKey = ""
                Exit Function
            End If
            Dim c As String = " "

            S = Right(c.PadRight(MaxKeyLen - 1, " ") + SecretKey + SecretKey + SecretKey + SecretKey, MaxKeyLen)
            X = Key2Num(Mid(EncodedKey, MaxKeyLen + 2, 1)) * CharCount + Key2Num(Mid(EncodedKey, MaxKeyLen + 1, 1))
            Data = ""
            For I = MaxKeyLen To 1 Step -1
                num = (CharCount + (Key2Num(Mid(EncodedKey, I, 1)) + 55 + Key2Num(Mid(S, I, 1)) - X) Mod CharCount) Mod CharCount
                t = Key2Char(num) + t
                X = X - 55 - Key2Num(Mid(S, I, 1)) + num
            Next I
            DecodeKey = Right(t, (CharCount * 100 + Key2Num(Right(EncodedKey, 1)) - X) Mod CharCount)
        Catch ex As Exception
            DecodeKey = ""
        End Try

    End Function

    Public Function Load_Permissions(ByVal Roles As String, ByRef usrPermissions As String) As Boolean
        Dim I As Integer
        Dim J As Integer
        Dim L As Integer
        Dim GID As Integer
        Dim Permission As String
        Dim P() As Integer
        'Dim DB As New DBUTIL

        ReDim P(255)

        For J = 1 To 255
            P(J) = 0
        Next J

        Load_Permissions = True
        usrPermissions = ""
        L = Len(Roles)
        For I = 1 To L
            GID = Asc(Mid$(Roles, I, 1)) - 64
            If GID = I Then
                Permission = DAL.GetSQLValue("SELECT RIGHTS FROM SYS_ROLES WHERE ROLE_ID=" & GID)
                If Permission <> "" Then
                    For J = 1 To Len(Permission)
                        P(J) = P(J) Or (Asc(Mid$(Permission, J, 1)) - 64)
                    Next J
                Else
                    Load_Permissions = False
                End If
            End If
        Next I

        For J = 1 To 255
            usrPermissions = usrPermissions + Chr(P(J) + 64)
        Next J
    End Function

#Region "User"
    ' Check if current user is in a specified role(s)
    ' i.e.  IsInRoles("Requester|Approver") 
    Public Function IsInRoles(ByVal Roles As String) As Boolean
        Dim role As String
        Dim Authorized As Boolean = False
        Dim S As String = ""

        If Not IsNothing(HttpContext.Current.Session("ROLES")) Then
            S = Strings.Join(HttpContext.Current.Session("ROLES"), "|")
        End If

        For Each role In Split(Roles, "|")
            'If Not HttpContext.Current.User.IsInRole(role) Then
            If (S & "").IndexOf(role) >= 0 Then
                Authorized = True
            End If
        Next

        Return (Authorized)
    End Function

    Public Sub DoCheckRole(ByVal PrivRole As String)
        If gReadConfig & "" = "Y" Then ReadConfigurations()
        If Not (IsAuthenticated() AndAlso Val(HttpContext.Current.Session("ROLES") & "") >= PrivRole) Then
            HttpContext.Current.Response.Redirect("../noRight.aspx")
        End If
    End Sub

    Public Sub CreateContext(ByVal UserName As String, ByVal Roles() As String)
        Dim identity As New System.Security.Principal.GenericIdentity(UserName)

        HttpContext.Current.User = New System.Security.Principal.GenericPrincipal(identity, Roles)
        'System.Threading.Thread.CurrentThread.CurrentPrincipal = New System.Security.Principal.GenericPrincipal(identity, Roles)
    End Sub

    Public Function IsAuthenticated() As Boolean
        Dim UserName As String = ""

        Try
            ' Check current logged on username from session variable first
            UserName = HttpContext.Current.Session("USER_NAME") & ""

            If UserName = "" Then
                ' Check from cookies
                If gUseCookies Then
                    UserName = HttpContext.Current.Request.Cookies("UserData").Values("code") & ""
                End If

                ' Check from Active Directory
                If UserName = "" AndAlso HttpContext.Current.User.Identity.IsAuthenticated Then
                    Dim I As Integer

                    UserName = HttpContext.Current.User.Identity.Name
                    I = UserName.IndexOf("\")
                    If I >= 0 Then UserName = UserName.Substring(I + 1)
                End If

                If UserName <> "" Then
                    LoadUserData(UserName, "", "IsAuthenticated")

                    UserName = HttpContext.Current.Session("USER_NAME") & ""

                    If UserName <> "" Then
                        'InsertLog(UserCode, "0", "1", "Log in")
                    End If
                End If
            End If
        Catch ex As Exception
            UserName = ""
        End Try

        Return (UserName <> "")
    End Function

    Public Function GetRoles(ByVal UserID As String) As String()
        Dim RET() As String = Nothing
        Dim Roles As String = ""
        Dim DR As DataRow

        DR = GetDR(DAL.SearchUserList(UserID, "", ""))
        If Not IsNothing(DR) Then
            Roles = DR("ROLE_ID") & "|"
            Return Roles.Split("|")
        Else
            Return Nothing
        End If
    End Function

    Public Function IsAppAuthenticated(ByVal UserName As String, Optional ByVal Password As String = "") As Boolean
        Dim DR As DataRow
        Dim RoleDesc(1) As String
        Dim DB As New DBUTIL

        Try
            'DR = GetDR(DAL.Login(UserName, Password))
            DR = GetDR(DAL.Login(UserName, ""))
            'Roles = GetRoles(UserName)
            If Not IsNothing(DR) Then
                If DR("PASSWORD") & "" = Password Then
                    If DR("DISABLED_FLAG") & "" = "Y" Then
                        IsAppAuthenticated = False
                    Else
                        If DR("PWD_EXPIRE_DATE") & "" <> "" AndAlso (CInt("0" & DR("DAY_EXPIRE")) > 0 And CDate(DR("PWD_EXPIRE_DATE")) < Today) Then
                            IsAppAuthenticated = False
                        End If
                        IsAppAuthenticated = True
                    End If
                Else
                    IsAppAuthenticated = False
                End If
            Else
                IsAppAuthenticated = False
            End If

        Catch ex As Exception
            IsAppAuthenticated = False
        End Try

        Return IsAppAuthenticated
    End Function

    Public Sub SetUserCookie(ByVal Key As String, ByVal Value As String)
        Dim Cookie As New HttpCookie("Tracking_User")

        Try
            Cookie.Values(Key) = Value
            Cookie.Expires = Now.Add(TimeSpan.FromMinutes(cCookieExpiration))
            HttpContext.Current.Response.Cookies.Add(Cookie)
        Catch
        End Try
    End Sub

    'Public Sub InitUserData(ByVal UserName As String, Optional ByVal ADPassword As String = "", Optional ByVal ADDomain As String = "")
    '    Dim LinktoPage As String = ""
    '    Dim Authenticated As Boolean = False
    '    Dim SU As New SecurityUtil
    '    Try
    '        If UserName <> "" Then
    '            ' Load personel data from WF_USERDETAIL$
    '            'If LoadPersonelData(UserName) Then
    '            '    'Default role as employee
    '            If LoadDefaultUserData(UserName, ADPassword, ADDomain) Then
    '                CreateSecurityContext(UserName, Split("1", "|"), "1")    ' employee
    '            End If

    '            ' Load user data from SYS_USERS (if exists)
    '            LoadAppUserData(UserName)
    '        End If
    '    Catch ex As Exception

    '    End Try

    'End Sub

    Public Sub CreatePrincipal(ByVal UserName As String, ByVal Roles() As String)
        Dim identity As New GenericIdentity(UserName)

        HttpContext.Current.User = New GenericPrincipal(identity, Roles)
        'CurrentThread.CurrentPrincipal = New GenericPrincipal(identity, Roles)
    End Sub

    Public Sub CreateSecurityContext(ByVal UserName As String, ByVal Roles() As String, Optional ByVal RoleID As String = "")

        HttpContext.Current.Session("USER_NAME") = UserName
        'HttpContext.Current.Session("USER_DESC") = UserName
        'HttpContext.Current.Session("ROLE_ID") = Roles
        HttpContext.Current.Session("ROLE_ID") = RoleID

        Try
            Dim xCookie As HttpCookie
            xCookie = New HttpCookie("UserData")
            xCookie.Values("UID") = UserName
            xCookie.Expires = DateTime.Now().Add(New TimeSpan(1, 2, 0, 0))  ' 2 hours
            HttpContext.Current.Response.Cookies.Add(xCookie)
            ClearObject(xCookie)

            CreatePrincipal(UserName, Roles)
        Catch
        End Try
    End Sub

    'Public Function IsAuthenticated() As Boolean
    '    Dim UserName As String = ""
    '    Dim Password As String = ""
    '    Dim I As Integer

    '    UserName = HttpContext.Current.Session("USER_NAME") & ""
    '    If UserName <> "" Then
    '        Return True
    '    Else
    '        Try
    '            'checked user from cookies
    '            If Not IsNothing(HttpContext.Current.Request.Cookies("UserData")) Then
    '                UserName = HttpContext.Current.Request.Cookies("UserData").Values("UID") & ""
    '            End If
    '        Catch ex As Exception
    '        End Try
    '    End If

    '    If (UserName = "") AndAlso HttpContext.Current.User.Identity.IsAuthenticated Then
    '        ' AD Authenticated
    '        UserName = HttpContext.Current.User.Identity.Name
    '        I = UserName.IndexOf("\")
    '        If I >= 0 Then UserName = UserName.Substring(I + 1)
    '    End If

    '    InitUserData(UserName.ToUpper)

    '    Return (HttpContext.Current.Session("USER_NAME") <> "")

    '    'If HttpContext.Current.User.Identity.IsAuthenticated Then
    '    '    If UserName = "" Then
    '    '        Dim I As Integer

    '    '        UserName = HttpContext.Current.User.Identity.Name
    '    '        I = UserName.IndexOf("\")
    '    '        If I >= 0 Then UserName = UserName.Substring(I + 1)
    '    '        If AdDomain <> "" Then LoadUserData(UserName)

    '    '        UserName = HttpContext.Current.Session("USER_NAME")
    '    '        If UserName <> "" Then
    '    '            Initialize()
    '    '            'WriteAppLog("Log On")
    '    '        End If
    '    '    End If
    '    'End If

    '    'Return (UserName <> "")
    'End Function

    Public Sub CheckSession(ByVal SData As String)
        'If SData = "" Then
        '    'HttpContext.Current.Response.Redirect("../Logout.aspx")
        '    HttpContext.Current.Response.Write("<script language=""javascript"">" & _
        '    "alert('Unable to access this page ! Please contact system administrator if you require to access.');" & _
        '    "if (opener){window.opener.location.href='../Logout.aspx'; this.close();}else if(parent)" & _
        '    "{window.parent.location.href='../Logout.aspx';} else" & _
        '    "{window.location.href='../Logout.aspx';}</script>")
        'End If
        Dim UID As String = Trim(HttpContext.Current.Session("UID") & "")
        If UID = "" OrElse UID <> HttpContext.Current.Request("UID") & "" Then
            If UID <> HttpContext.Current.Request("UID") & "" Then
                If UID <> "" Then
                    BLL.InsertAudit(catBannedRequestLog, "Too many request failure!", HttpContext.Current.Session("USER_NAME") & "")
                End If
                ClearSession()
            End If
            HttpContext.Current.Response.Write("<script language=""javascript"">if (parent != null || parent != undefined)" & _
            "{parent.location.href='../NoRight.htm';} else if (opener != null || opener != undefined)" & _
            "{opener.location.href='../NoRight.htm'; this.close();} else {window.location.href='../NoRight.htm';}</script>")
        End If
    End Sub

    Public Sub ClearSession()
        Dim I As Integer

        Try
            ' Clear session variables
            With HttpContext.Current.Session
                For I = 0 To .Count - 1
                    ClearObject(.Item(I))
                Next

                .Clear()
                .Abandon()
            End With

            ' Clear cookies
            If gUseCookies Then SetUserCookie("code", "")

            If Not IsNothing(HttpContext.Current.Response.Cookies("ASP.NET_SessionId")) Then
                HttpContext.Current.Response.Cookies("ASP.NET_SessionId").Expires = DateTime.Now.AddYears(-30)
            End If
        Catch
        End Try
    End Sub

    Public Sub InitUserData(ByVal UserName As String, Optional ByVal ADPassword As String = "", Optional ByVal ADDomain As String = "")
        Dim LinktoPage As String = ""
        Dim Authenticated As Boolean = False
        Dim SU As New SecurityUtil
        Try
            If UserName <> "" Then

                'If LoadDefaultUserData(UserName, ADPassword, ADDomain) Then
                '    CreateSecurityContext(UserName, Split("1", "|"), "1")    ' employee
                'End If

                ' Load user data from SYS_USERS (if exists)
                LoadUserData(UserName)
            End If
        Catch ex As Exception

        End Try

    End Sub

    ' Check banned log
    Public Function IsBanned(ByVal UserName As String) As Boolean
        Dim DR As DataRow
        Dim m As Double
        Dim BannedFlag As String = ""
        Try
            DR = GetDR(DAL.SearchAudit("", "", catBannedLog, "", UserName.ToUpper, OrderSQL:="SL.TRANS_DATE DESC"))
            m = DateDiff(DateInterval.Minute, DR("TRANS_DATE"), Now)
            BannedFlag = DAL.LookupSQL("SELECT BANNED_FLAG FROM SYS_USERS WHERE USER_NAME='" & UserName.ToUpper & "'")
            If (Not IsNothing(DR) AndAlso m < ToInt(gBannedDuration)) AndAlso BannedFlag = "Y" Then
                Return True
            Else
                Return False
            End If
        Catch ex2 As Exception
            ' Ignore error
        End Try

    End Function

    ' Check banned log
    Public Function IsBannedRequest(ByVal UserName As String) As Boolean
        Dim DT As DataTable = Nothing
        Dim DR As DataRow = Nothing
        Dim m As Double
        Dim BannedFlag As String = ""
        Try
            DT = DAL.SearchAudit("", "", catBannedRequestLog, "", UserName.ToUpper, OtherCriteria:="FLOOR((SYSDATE-SL.TRANS_DATE)*24) <= 2", OrderSQL:="SL.TRANS_DATE DESC")
            DR = GetDR(DT)
            If DT.Rows.Count > 3 Then
                m = DateDiff(DateInterval.Minute, DR("TRANS_DATE"), Now)
                'BannedFlag = DAL.LookupSQL("SELECT BANNED_FLAG FROM SYS_USERS WHERE USER_NAME='" & UserName.ToUpper & "'")
                If (Not IsNothing(DR) AndAlso m < ToInt(gBannedDuration)) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex2 As Exception
            ' Ignore error
        Finally
            ClearObject(DR) : ClearObject(DT)
        End Try

    End Function

    'Public Sub PreventPostbackCSRF(ByVal EnableViewStatMax As Boolean, ByVal IsPostBack As Boolean)
    '    If Not EnableViewStatMax Then

    '    End If
    'End Sub
#End Region
End Module




