#Region ".NET Framework Class Import"
Imports System.Data
Imports System.Web
#End Region


Partial Public Class DALComponent

#Region "Master Data"
    Public Function SearchTitle(Optional ByVal TitleID As String = "", Optional ByVal TitleNameE As String = "" _
    , Optional ByVal TitleNameT As String = "", Optional ByVal OtherCriteria As String = "" _
    , Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "TITLE_ID", TitleID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "TITLE_NAME_E", TitleNameE, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "TITLE_NAME_T", TitleNameT, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_TITLES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY TITLE_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchGroup(Optional ByVal UserGroupID As String = "", Optional ByVal UserGroupName As String = "", Optional ByVal RoleID As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UG.GROUP_ID", UserGroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UG.ROLE_ID", RoleID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(UG.GROUP_NAME)", UserGroupName.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT UG.*,SR.ROLE_NAME FROM SYS_GROUPS UG,SYS_ROLES SR WHERE UG.ROLE_ID=SR.ROLE_ID(+) "
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY UG.GROUP_NAME"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchUserLevel(Optional ByVal UserLevel As String = "", Optional ByVal UserLevelName As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "LEVEL_ID", UserLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(LEVEL_NAME)", UserLevelName.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_USER_LEVELS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY LEVEL_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchServiceStatus(Optional ByVal ServiceStatus As String = "", Optional ByVal ServiceStatusDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SERVICE_STATUS_DESC)", ServiceStatusDesc.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_SERVICE_STATUS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SERVICE_STATUS"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchServiceType(Optional ByVal ServiceType As String = "", Optional ByVal ServiceTypeDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SERVICE_TYPE_DESC)", ServiceTypeDesc.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_SERVICE_TYPES"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SERVICE_TYPE"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSiteStatus(Optional ByVal SiteStatus As String = "", Optional ByVal SiteStatusDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SITE_STATUS", SiteStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SITE_STATUS_DESC)", SiteStatusDesc.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM REF_SITE_STATUS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SITE_STATUS"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchProblemResolve(Optional ByVal ProblemResolveID As String = "", Optional ByVal ProblemType As String = "" _
    , Optional ByVal ProblemResolveDesc As String = "", Optional ByVal ProjectType As String = "", Optional ByVal Detail As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PR.PROB_RESOLVE_ID", ProblemResolveID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PR.PROBLEM_TYPE", ProblemType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(PR.PROB_RESOLVE_DESC)", ProblemResolveDesc.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PR.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)

            DB.AddCriteria(Criteria2, "UPPER(PROB_RESOLVE_DETAIL_DESC)", Detail, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT PR.*,PT.PROBLEM_TYPE_DESC FROM PROBLEM_RESOLVES PR,REF_PROBLEM_TYPES PT WHERE PR.PROBLEM_TYPE=PT.PROBLEM_TYPE(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If Criteria2 <> "" Then SQL &= " AND PR.PROB_RESOLVE_ID IN (SELECT PROB_RESOLVE_ID FROM PROBLEM_RESOLVE_DETAILS WHERE " & Criteria2 & ")"
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PR.PROB_RESOLVE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchProblemType(Optional ByVal ProblemType As String = "", Optional ByVal ProblemTypeDesc As String = "" _
    , Optional ByVal ProjectType As String = "", Optional ByVal OtherCriteria As String = "" _
    , Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PT.PROBLEM_TYPE", ProblemType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PT.PROBLEM_TYPE_DESC", ProblemTypeDesc, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PT.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT PT.*,PJ.PROJECT_TYPE_DESC,PJ.SERVICE_PREFIX_NAME " & _
            " FROM REF_PROBLEM_TYPES PT,REF_PROJECT_TYPES PJ WHERE PT.PROJECT_TYPE=PJ.PROJECT_TYPE(+) "
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PT.PROBLEM_TYPE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchProblemResolveDetail(Optional ByVal ProbResolveID As String = "", Optional ByVal ProbResolveDetailID As String = "" _
    , Optional ByVal ProbResolveDetailDesc As String = "", Optional ByVal OtherCriteria As String = "" _
    , Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PROB_RESOLVE_DETAIL_ID", ProbResolveDetailID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(PROB_RESOLVE_DETAIL_DESC)", ProbResolveDetailDesc.ToUpper, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM PROBLEM_RESOLVE_DETAILS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY PROB_RESOLVE_DETAIL_ID"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSeverityLevel(Optional ByVal SeverityLevel As String = "", Optional ByVal SeverityLevelDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SEVERITY_LEVEL_DESC", SeverityLevelDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_SEVERITY_LEVELS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SEVERITY_LEVEL_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSLADetail(Optional ByVal SLAProfileID As String = "", Optional ByVal SeverityLevel As String = "" _
    , Optional ByVal ProjectType As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SD.SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SD.SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SP.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)


            SQL = "SELECT SD.*,SL.SEVERITY_LEVEL_DESC FROM V_SLA_DETAILS SD,SLA_PROFILES SP,REF_SEVERITY_LEVELS SL " & _
            "WHERE SD.SEVERITY_LEVEL=SL.SEVERITY_LEVEL(+) AND SD.SLA_PROFILE_ID=SP.SLA_PROFILE_ID(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SD.SLA_PROFILE_ID,SD.SEVERITY_LEVEL"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchVendorStaff(Optional ByVal StaffID As String = "", Optional ByVal VendorCode As String = "" _
    , Optional ByVal ProjectType As String = "", Optional ByVal StaffName As String = "" _
    , Optional ByVal StaffCode As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "STAFF_ID", StaffID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(VENDOR_CODE)", VendorCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(STAFF_CODE)", StaffCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(STAFF_NAME)", StaffName.ToUpper, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM VENDOR_STAFFS "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEquipmentType(Optional ByVal EquipType As String = "", Optional ByVal EquipTypeDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQUIPMENT_TYPE_DESC", EquipTypeDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_EQUIPMENT_TYPES "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EQUIPMENT_TYPE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEquipmentStatus(Optional ByVal EquipStatus As String = "", Optional ByVal EquipStatusDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "EQUIPMENT_STATUS", EquipStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQUIPMENT_STATUS_DESC", EquipStatusDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_EQUIPMENT_STATUS "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EQUIPMENT_STATUS_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchSystemStatus(Optional ByVal SystemStatus As String = "", Optional ByVal SystemStatusDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SYSTEM_STATUS", SystemStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SYSTEM_STATUS_DESC", SystemStatusDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_SYSTEM_STATUS "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SYSTEM_STATUS"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function SearchMovementType(Optional ByVal MovementType As String = "", Optional ByVal MovementTypeDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "MOVEMENT_TYPE", MovementType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(MOVEMENT_TYPE_DESC)", MovementTypeDesc.ToUpper, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_MOVEMENT_TYPES "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY MOVEMENT_TYPE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function



    Public Function SearchBOMGroup() As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            SQL = "SELECT BOM_GROUP_ID,NAME From Service_BOM_GROUP "

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchBOMDetail() As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            SQL = "SELECT BOM_DETAIL_ID,DESCRIPTION From Service_BOM_DETAIL"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Public Function SearchSystemStatus() As DataTable
    '    Dim DT As DataTable = Nothing
    '    Dim SQL As String = "", Criteria As String = ""

    '    Try
    '        SQL = "SELECT System_Status,System_Status_desc From REF_System_Status"

    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Public Function SearchUnit(Optional ByVal UnitID As String = "", Optional ByVal UnitDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UNIT_ID", UnitID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UNIT_DESC", UnitDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_UNITS "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY UNIT_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchWarrantyType(Optional ByVal WarrantyType As String = "", Optional ByVal WarrantyTypeDesc As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "WARRANTY_TYPE", WarrantyType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "WARRANTY_TYPE_DESC", WarrantyTypeDesc, DBUTIL.FieldTypes.ftText)


            SQL = "SELECT * FROM REF_WARRANTY_TYPES "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY WARRANTY_TYPE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageSLAProfile(ByVal op As Integer, Optional ByRef SLAProfileID As String = Nothing _
    , Optional ByVal ProfileName As String = Nothing, Optional ByVal ProjectType As String = Nothing _
    , Optional ByVal VendorCode As String = Nothing, Optional ByVal ActiveFlag As String = Nothing _
    , Optional ByVal SLAType As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If SLAProfileID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    SLAProfileID = GenerateID("SLA_PROFILES", "SLA_PROFILE_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "PROFILE_NAME", ProfileName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SLA_TYPE", SLAType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ACTIVE_FLAG", ActiveFlag, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SLA_PROFILES", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                SLAProfileID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function ManageSiteGroup(ByVal op As Integer, Optional ByRef SiteGrpID As String = Nothing _
    , Optional ByVal SiteGrpName As String = Nothing, Optional ByVal Province As String = Nothing _
    , Optional ByVal Remark As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SITE_GROUP_ID", SiteGrpID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If SiteGrpID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "SITE_GROUP_ID", SiteGrpID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    SiteGrpID = GenerateID("SITE_GROUPS", "SITE_GROUP_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "SITE_GROUP_ID", SiteGrpID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "SITE_GROUP_NAME", SiteGrpName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROVINCE_ID", Province, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "REMARK", Remark, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SITE_GROUPS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                SiteGrpID = ""
            End If
            Throw ex
        End Try

    End Function

    'Public Function ManageSLADetail(ByVal op As Integer, Optional ByRef SLAProfileID As String = Nothing _
    ', Optional ByVal SeverityLevel As String = Nothing, Optional ByVal ResponseDay As String = Nothing _
    ', Optional ByVal ResponseHour As String = Nothing, Optional ByVal ResponseMinute As String = Nothing _
    ', Optional ByVal ResolutionDay As String = Nothing, Optional ByVal ResolutionHour As String = Nothing _
    ', Optional ByVal ResolutionMinute As String = Nothing) As String

    '    Dim SQL1, SQL2, SQL As String
    '    Dim Criteria As String = ""

    '    Try
    '        SQL = "" : SQL1 = "" : SQL2 = ""

    '        If op = DBUTIL.opDELETE Then
    '            DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '        Else
    '            If op = DBUTIL.opUPDATE Then
    '                DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '            Else
    '                DB.AddSQL(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddSQL(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '            End If
    '        End If

    '        DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_DAY", ResponseDay, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_HOUR", ResponseHour, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_MINUTE", ResponseMinute, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL2(op, SQL1, SQL2, "RESOLUTION_DAY", ResolutionDay, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL2(op, SQL1, SQL2, "RESOLUTION_HOUR", ResolutionHour, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL2(op, SQL1, SQL2, "RESOLUTION_MINUTE", ResolutionMinute, DBUTIL.FieldTypes.ftNumeric)
    '        SQL = DB.CombineSQL(op, SQL1, SQL2, "SLA_DETAILS", Criteria, True)
    '        DB.ExecSQL(SQL)
    '        Return ""
    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    Public Function ManageSLADetail(ByVal op As Integer, Optional ByRef SLAProfileID As String = Nothing _
    , Optional ByRef SeverityLevelOld As String = Nothing, Optional ByVal SeverityLevel As String = Nothing _
    , Optional ByVal ResponseTime As String = Nothing _
    , Optional ByVal ResolutionTime As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevelOld, DBUTIL.FieldTypes.ftNumeric)
            Else
                If op = DBUTIL.opUPDATE Then
                    DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevelOld, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
                Else
                    DB.AddSQL(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
                End If
            End If

            DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_TIME", ResponseTime, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL2(op, SQL1, SQL2, "RESOLUTION_TIME", ResolutionTime, DBUTIL.FieldTypes.ftNumeric)
            SQL = DB.CombineSQL(op, SQL1, SQL2, "SLA_DETAILS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function ManageVendor(ByVal op As Integer, Optional ByRef VendorCode As String = Nothing _
    , Optional ByVal VendorName As String = Nothing, Optional ByVal ProjectType As String = Nothing _
    , Optional ByVal Address As String = Nothing, Optional ByVal ZipCode As String = Nothing _
    , Optional ByVal ProvinceID As String = Nothing, Optional ByVal TelNo As String = Nothing _
    , Optional ByVal FaxNo As String = Nothing, Optional ByVal VendorCodeSAP As String = Nothing _
    , Optional ByVal Email As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    DB.AddSQL(op, SQL1, SQL2, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
                Else
                    op = DBUTIL.opUPDATE
                End If

                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_NAME", VendorName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ADDRESS", Address, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROVINCE_ID", ProvinceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ZIP_CODE", ZipCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "FAX_NO", FaxNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_CODE_SAP", VendorCodeSAP, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EMAIL", Email, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "VENDORS", Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function ManageVendorStaff(ByVal op As Integer, Optional ByRef StaffID As String = Nothing _
    , Optional ByVal VendorCode As String = Nothing, Optional ByVal StaffName As String = Nothing _
    , Optional ByVal TelNo As String = Nothing, Optional ByVal MobileNo As String = Nothing _
    , Optional ByVal Email As String = Nothing, Optional ByVal StaffCode As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "STAFF_ID", StaffID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    StaffID = GenerateID("VENDOR_STAFFS", "STAFF_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "STAFF_ID", StaffID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
                Else
                    op = DBUTIL.opUPDATE
                End If

                DB.AddSQL2(op, SQL1, SQL2, "STAFF_NAME", StaffName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "MOBILE_NO", MobileNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EMAIL", Email, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "STAFF_CODE", StaffCode, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "VENDOR_STAFFS", Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function ManageProblemResolve(ByVal op As Integer, Optional ByRef ProbResolveID As String = Nothing _
    , Optional ByVal ProblemType As String = Nothing, Optional ByVal ProbResolveDesc As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If ProbResolveID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    ProbResolveID = GenerateID("PROBLEM_RESOLVES", "PROB_RESOLVE_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "PROB_RESOLVE_DESC", ProbResolveDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM_TYPE", ProblemType, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "PROBLEM_RESOLVES", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                ProbResolveID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function ManageProblemResolveDetail(ByVal op As Integer, Optional ByVal ProbResolveID As String = Nothing _
    , Optional ByRef ProbResolveDetailID As String = Nothing, Optional ByVal ProbResolveDetailDesc As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> DBUTIL.opINSERT Then
                DB.AddCriteria(Criteria, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "PROB_RESOLVE_DETAIL_ID", ProbResolveDetailID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> DBUTIL.opDELETE Then
                If op = DBUTIL.opINSERT Then
                    op = DBUTIL.opINSERT
                    ProbResolveDetailID = GenerateID("PROBLEM_RESOLVE_DETAILS", "PROB_RESOLVE_DETAIL_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "PROB_RESOLVE_ID", ProbResolveID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "PROB_RESOLVE_DETAIL_ID", ProbResolveDetailID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opUPDATE
                End If

                DB.AddSQL2(op, SQL1, SQL2, "PROB_RESOLVE_DETAIL_DESC", ProbResolveDetailDesc, DBUTIL.FieldTypes.ftText)
            End If

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "PROBLEM_RESOLVE_DETAILS", Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region "Config Table"
    'Created By Aoy 18/06/2552
    Public Function SearchConfigMenu(Optional ByVal MenuID As String = "" _
    , Optional ByVal ParentMenuID As String = "", Optional ByVal Active As String = "" _
    , Optional ByVal TableID As String = "", Optional ByVal TableName As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "MENU_ID", MenuID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PARENT_MENU_ID", ParentMenuID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "ACTIVE_FLAG", Active, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "TABLE_CODE", TableID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "TABLE_NAME", TableName, DBUTIL.FieldTypes.ftText)
            SQL = "SELECT * FROM CONFIG_MENUS"
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY ORDER_NO"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 03/08/2552 Updated By Aoy 20/11/2552
    Public Function SearchConfigTables(Optional ByVal TableID As String = "" _
    , Optional ByVal TableName As String = "", Optional ByVal OtherCriteria As String = "" _
    , Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria2(Criteria, "UPPER(CT.TABLE_CODE)", TableID.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CT.TABLE_NAME)", TableName.ToUpper(), DBUTIL.FieldTypes.ftText)
            SQL = "SELECT CT.*,CM.MENU_DESC_T,CM.MENU_DESC_E FROM CONFIG_TABLES CT, " & _
            "CONFIG_MENUS CM WHERE CT.TABLE_CODE=CM.TABLE_CODE(+) "
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY TABLE_NAME"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 18/06/2552 Updated By Aoy 20/11/2552
    Public Function SearchConfigTableDetails(Optional ByVal TableID As String = "" _
    , Optional ByVal ColumnName As String = "", Optional ByVal ActiveFlag As String = "" _
    , Optional ByVal KeyFlag As String = "", Optional ByVal IsFKDisplayFlag As String = "" _
    , Optional ByVal DisplayFlag As String = "", Optional ByVal IsSearchFlag As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria2(Criteria, "UPPER(CTD.TABLE_CODE)", TableID.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.COLUMN_NAME)", ColumnName.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.ACTIVE_FLAG)", ActiveFlag.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.DISPLAY_FLAG)", DisplayFlag.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.KEY_FLAG)", KeyFlag.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.IS_FK_DISPLAY_FLAG)", IsFKDisplayFlag.ToUpper(), DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2(Criteria, "UPPER(CTD.SEARCH_FLAG)", IsSearchFlag.ToUpper(), DBUTIL.FieldTypes.ftText)
            SQL = "SELECT CTD.*,CT.TABLE_NAME,CT.DISPLAY_FLAG,CCT.CTRL_TYPE " & _
            " FROM CONFIG_TABLE_DETAILS CTD " & _
            ",CONFIG_TABLES CT,CONFIG_COLUMN_TYPES CCT WHERE CTD.TABLE_CODE=CT.TABLE_CODE(+) " & _
            " AND CTD.COLUMN_TYPE = CCT.COLUMN_TYPE(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            SQL &= " ORDER BY CTD.ORDER_NO"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 18/06/2552
    Public Function SearchConfigColumnTypes(Optional ByVal ColumnType As String = "", Optional ByVal OtherCriteria As String = "" _
    , Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria2(Criteria, "UPPER(COLUMN_TYPE)", ColumnType.ToUpper, DBUTIL.FieldTypes.ftText)
            SQL = "SELECT *  " & _
            " FROM CONFIG_COLUMN_TYPES "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            SQL &= " ORDER BY DATE_UPDATED DESC"

            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 18/06/2552 Updated By Aoy 20/11/2552
    Public Function SearchMasterData(ByVal TableID As String, ByVal TableName As String _
    , Optional ByVal SearchValue As String = "", Optional ByVal IsSelectAll As Boolean = True _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim DT2 As DataTable = Nothing
        Dim DT3 As DataTable = Nothing
        Dim DT4 As DataTable = Nothing
        Dim DR2, DR3 As DataRow
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = "", ColumnList As String = "" _
        , TableList As String = ""
        Dim AbbrTable(20) As String
        Dim TableCnt As Integer = 1
        Dim arrSearchValue As String()
        Try
            Criteria = OtherCriteria
            DT3 = SearchConfigTableDetails(TableID)
            arrSearchValue = SearchValue.Split(",")
            For I As Integer = 0 To arrSearchValue.Count - 1
                If arrSearchValue(I) <> "" Then
                    GenerateCriteria(Criteria, TableName & "." & DT3.Rows(I)("COLUMN_NAME") & "", DT3.Rows(I)("DATA_TYPE") & "", arrSearchValue(I))
                End If
            Next

            TableList = TableName

            If IsSelectAll Then
                ColumnList = TableName & ".*,"
                DT2 = SearchConfigTableDetails(TableID, OtherCriteria:="(CTD.REF_TABLE_CODE IS NOT NULL)")
            Else
                DT2 = SearchConfigTableDetails(TableID)
            End If

            For Each DR2 In DT2.Rows
                If DR2("REF_TABLE_CODE") & "" <> "" Then
                    DT3 = SearchConfigTableDetails(DR2("REF_TABLE_CODE") & "", IsFKDisplayFlag:="Y")
                    If Not IsNothing(DT3) Then
                        DR3 = GetDR(DT3)
                        If Not IsNothing(DR3) Then
                            TableList &= " ," & DR3("TABLE_NAME") & ""
                            If Criteria2 <> "" Then Criteria2 &= " AND "
                            Criteria2 &= TableName & "." & DR2("COLUMN_NAME") & ""

                            If DR2("REF_PARAM1") & "" <> "" Then
                                Criteria2 &= " = " & DR3("TABLE_NAME") & "." & DR2("REF_PARAM1") & "(+)"
                            Else
                                Criteria2 &= " = " & DR3("TABLE_NAME") & "." & DR2("COLUMN_NAME") & "(+)"
                            End If
                            For Each DR3 In DT3.Rows
                                ColumnList &= DR3("TABLE_NAME") & "." & DR3("COLUMN_NAME") & ","
                            Next
                        End If
                    End If
                Else
                    ColumnList &= TableName & "." & DR2("COLUMN_NAME") & ","
                End If
                If OrderBy = "" And DR2("SORT_FLAG") & "" = "Y" Then
                    If OrderBy <> "" Then OrderBy &= ","
                    OrderBy &= TableName & "." & DR2("COLUMN_NAME") & ""
                End If
            Next

            If ColumnList <> "" Then
                ColumnList = Left(ColumnList, ColumnList.Length - 1)
            End If

            If Criteria2 <> "" Then
                If Criteria <> "" Then
                    Criteria = Criteria2 & " AND (" & Criteria & ")"
                Else
                    Criteria = Criteria2
                End If
            End If

            SQL = "SELECT " & ColumnList & " FROM " & TableList
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY " & TableName & ".DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        Finally
            ClearObject(DT3)
            ClearObject(DT2)
            ClearObject(DT)
        End Try
    End Function

    'Public Function SearchMasterData(ByVal TableID As String, ByVal TableName As String _
    ', Optional ByVal SearchValue As String = "", Optional ByVal IsSelectAll As Boolean = True _
    ', Optional ByVal OtherCriteria As String = "" , Optional ByVal OrderBy As String = "") As DataTable
    '    Dim DT As DataTable = Nothing
    '    Dim DT2 As DataTable = Nothing
    '    Dim DT3 As DataTable = Nothing
    '    Dim DT4 As DataTable = Nothing
    '    Dim DR2, DR3 As DataRow
    '    Dim SQL As String = "", Criteria As String = "", ColumnList As String = "" _
    '    , TableList As String = ""
    '    Dim AbbrTable(20) As String
    '    Dim TableCnt As Integer = 1
    '    Dim arrSearchValue As String()
    '    Try
    '        Criteria = OtherCriteria
    '        DT3 = SearchConfigTableDetails(TableID)
    '        arrSearchValue = SearchValue.Split(",")
    '        For I As Integer = 0 To arrSearchValue.Count - 1
    '            If arrSearchValue(I) <> "" Then
    '                GenerateCriteria(Criteria, "T1." & DT3.Rows(I)("COLUMN_NAME") & "", DT3.Rows(I)("DATA_TYPE") & "", arrSearchValue(I))
    '            End If
    '        Next
    '        AbbrTable(TableCnt - 1) = "T" & TableCnt & ""
    '        TableList = TableName & " AS " & AbbrTable(TableCnt - 1)

    '        If IsSelectAll Then
    '            ColumnList = AbbrTable(0) & ".*,"
    '            'DT2 = SearchConfigTableDetails(TableID, OtherCriteria:="(CTD.KEY_FLAG='FK' OR CTD.REF_TABLE_CODE IS NOT NULL)")
    '            DT2 = SearchConfigTableDetails(TableID, OtherCriteria:="(CTD.REF_TABLE_CODE IS NOT NULL)")
    '        Else
    '            DT2 = SearchConfigTableDetails(TableID)
    '        End If

    '        For Each DR2 In DT2.Rows
    '            'If DR2("KEY_FLAG") & "" = "FK" OrElse DR2("REF_TABLE_CODE") & "" <> "" Then
    '            If DR2("REF_TABLE_CODE") & "" <> "" Then
    '                TableCnt = TableCnt + 1
    '                AbbrTable(TableCnt - 1) = "T" & TableCnt & ""
    '                DT3 = SearchConfigTableDetails(DR2("REF_TABLE_CODE") & "", IsFKDisplayFlag:="Y")
    '                If Not IsNothing(DT3) Then
    '                    DR3 = GetDR(DT3)
    '                    If Not IsNothing(DR3) Then
    '                        TableList &= " LEFT OUTER JOIN " & DR3("TABLE_NAME") & " AS " & _
    '                        AbbrTable(TableCnt - 1) & " ON " & AbbrTable(0) & "." & DR2("COLUMN_NAME") & ""

    '                        If DR2("REF_PARAM1") & "" <> "" Then
    '                            TableList &= " = " & AbbrTable(TableCnt - 1) & "." & DR2("REF_PARAM1") & ""
    '                        Else
    '                            TableList &= " = " & AbbrTable(TableCnt - 1) & "." & DR2("COLUMN_NAME") & ""
    '                        End If
    '                        For Each DR3 In DT3.Rows
    '                            ColumnList &= AbbrTable(TableCnt - 1) & "." & DR3("COLUMN_NAME") & ","
    '                        Next

    '                    End If
    '                End If
    '            Else
    '                ColumnList &= AbbrTable(0) & "." & DR2("COLUMN_NAME") & ","
    '            End If
    '            If OrderBy = "" And DR2("SORT_FLAG") & "" = "Y" Then
    '                If OrderBy <> "" Then OrderBy &= ","
    '                OrderBy &= AbbrTable(0) & "." & DR2("COLUMN_NAME") & ""
    '            End If
    '        Next

    '        If ColumnList <> "" Then
    '            ColumnList = Left(ColumnList, ColumnList.Length - 1)
    '        End If

    '        SQL = "SELECT " & ColumnList & " FROM " & TableList
    '        If Criteria <> "" Then SQL &= " WHERE " & Criteria
    '        If OrderBy <> "" Then SQL &= " ORDER BY " & OrderBy
    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ClearObject(DT3)
    '        ClearObject(DT2)
    '        ClearObject(DT)
    '    End Try
    'End Function

    Public Sub GenerateCriteria(ByRef Criteria As String, ByVal ColumnName As String _
    , ByVal DataType As String, ByVal Value1 As String, Optional ByVal Value2 As String = "")
        Try
            If Value1 <> "" And Value2 <> "" Then
                Select Case DataType
                    Case "TEXT" : DB.AddCriteriaRange(Criteria, "UPPER(" & ColumnName & ")", Value1.ToUpper, Value2.ToUpper, DBUTIL.FieldTypes.ftText)
                    Case "NUMERIC" : DB.AddCriteriaRange(Criteria, ColumnName, Value1, Value2, DBUTIL.FieldTypes.ftNumeric)
                    Case "DATE" : DB.AddCriteriaRange(Criteria, ColumnName, AppDateValue(Value1), AppDateValue(Value2), DBUTIL.FieldTypes.ftDate)
                    Case "DATETIME" : DB.AddCriteriaRange(Criteria, ColumnName, AppDateValue(Value1), AppDateValue(Value2), DBUTIL.FieldTypes.ftDateTime)
                    Case "BINARY" : DB.AddCriteriaRange(Criteria, ColumnName, Value1, Value2, DBUTIL.FieldTypes.ftBinary)
                    Case Else
                        DB.AddCriteriaRange(Criteria, ColumnName, Value1, Value2, DBUTIL.FieldTypes.ftText)
                End Select
            Else
                Select Case DataType
                    Case "TEXT" : DB.AddCriteria(Criteria, "UPPER(" & ColumnName & ")", Value1.ToUpper, DBUTIL.FieldTypes.ftText)
                    Case "NUMERIC" : DB.AddCriteria(Criteria, ColumnName, Value1, DBUTIL.FieldTypes.ftNumeric)
                    Case "DATE" : DB.AddCriteria(Criteria, ColumnName, AppDateValue(Value1), DBUTIL.FieldTypes.ftDate)
                    Case "DATETIME" : DB.AddCriteria(Criteria, ColumnName, AppDateValue(Value1), DBUTIL.FieldTypes.ftDateTime)
                    Case "BINARY" : DB.AddCriteria(Criteria, ColumnName, Value1, DBUTIL.FieldTypes.ftBinary)
                    Case Else
                        DB.AddCriteria(Criteria, "UPPER(" & ColumnName & ")", Value1.ToUpper, DBUTIL.FieldTypes.ftText)
                End Select
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function GenerateSQLValue(ByVal DataType As String _
    , ByVal Value As String) As String
        Dim ret As String = ""
        Try
            Select Case DataType
                Case "TEXT" : ret = DB.SQLValue(Value, DBUTIL.FieldTypes.ftText).ToString().Trim()
                Case "NUMERIC" : ret = DB.SQLValue(Value, DBUTIL.FieldTypes.ftNumeric).ToString().Trim()
                Case "DATE" : ret = DB.SQLValue(AppDateValue(Value), DBUTIL.FieldTypes.ftDate).ToString().Trim()
                Case "DATETIME" : ret = DB.SQLValue(AppDateValue(Value), DBUTIL.FieldTypes.ftDateTime).ToString().Trim()
                Case "BINARY" : ret = DB.SQLValue(Value, DBUTIL.FieldTypes.ftBinary).ToString().Trim()
                Case Else
                    ret = DB.SQLValue(Value, DBUTIL.FieldTypes.ftText).ToString().Trim()
            End Select
        Catch ex As Exception
            ret = ""
        End Try
        Return ret
    End Function

    'Updated By Aoy 22/06/2552
    Public Function ManageMasterData(ByVal op As Integer, ByVal TableName As String, ByVal ArrList As ArrayList _
                                     , ByVal DTConfig As DataTable, Optional ByVal OldValue As String = "", Optional ByVal ReadOnlyFlag As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""
        Dim DR As DataRow
        Dim I As Integer = 0
        Dim arrOldValue As String()
        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            arrOldValue = OldValue.Split(",")
            For Each DR In DTConfig.Rows
                If op <> DBUTIL.opINSERT Then
                    If DR("KEY_FLAG") & "" = "PK" Then
                        DB.AddCriteria(Criteria, DR("COLUMN_NAME") & "", arrOldValue(I), GetDBFieldType(DR("DATA_TYPE") & ""))
                    End If
                End If
                If op <> DBUTIL.opDELETE Then
                    If op = DBUTIL.opINSERT Then
                        op = DBUTIL.opINSERT
                        If DR("KEY_FLAG") & "" = "PK" AndAlso DR("EDIT_FLAG") & "" <> "N" Then
                            DB.AddSQL2(op, SQL1, SQL2, DR("COLUMN_NAME") & "", ArrList.Item(I), GetDBFieldType(DR("DATA_TYPE") & ""))
                        End If
                        If ReadOnlyFlag = "N" Then
                            DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                            DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
                        End If
                    Else
                        op = DBUTIL.opUPDATE
                        If DR("KEY_FLAG") & "" = "PK" AndAlso DR("EDIT_FLAG") & "" <> "N" Then
                            DB.AddSQL2(op, SQL1, SQL2, DR("COLUMN_NAME") & "", ArrList.Item(I), GetDBFieldType(DR("DATA_TYPE") & ""))
                        End If
                    End If

                    If DR("KEY_FLAG") & "" <> "PK" AndAlso DR("EDIT_FLAG") & "" <> "N" Then
                        DB.AddSQL2(op, SQL1, SQL2, DR("COLUMN_NAME") & "", ArrList.Item(I), GetDBFieldType(DR("DATA_TYPE") & ""))
                    End If
                End If
                I = I + 1
            Next

            If op <> DBUTIL.opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, TableName, Criteria, True)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Equipment"
    Public Function SearchEquipment(Optional ByVal EquipID As String = "" _
    , Optional ByVal SerialNo As String = "", Optional ByVal EquipName As String = "", Optional ByVal EquipShortName As String = "" _
    , Optional ByVal EquipType As String = "", Optional ByVal Brand As String = "", Optional ByVal Model As String = "" _
    , Optional ByVal InstallDataF As String = "", Optional ByVal InstallDateT As String = "" _
    , Optional ByVal EquipStatus As String = "", Optional ByVal SiteName As String = "", Optional ByVal EquipSet As String = "" _
    , Optional ByVal SiteID As String = "", Optional ByVal BarCodeNo As String = "", Optional ByVal SAPNo As String = "" _
    , Optional ByVal PartNo As String = "", Optional ByVal RDNumber As String = "", Optional ByVal KioskNo As String = "" _
    , Optional ByVal Version As String = "", Optional ByVal QuantityF As String = "", Optional ByVal QuantityT As String = "" _
    , Optional ByVal Unit As String = "", Optional ByVal EquipOwner As String = "", Optional ByVal CostPerUnitF As String = "" _
    , Optional ByVal CostPerUnitT As String = "", Optional ByVal WADateF As String = "", Optional ByVal WADateT As String = "" _
    , Optional ByVal WAType As String = "", Optional ByVal PMDateF As String = "", Optional ByVal PMDateT As String = "" _
    , Optional ByVal UpgradeDateF As String = "", Optional ByVal UpgradeDateT As String = "", Optional ByVal Location As String = "" _
    , Optional ByVal UseAgeF As String = "", Optional ByVal UseAgeT As String = "", Optional ByVal Network As String = "" _
    , Optional ByVal System As String = "", Optional ByVal Detail As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "EQ.EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EQ.SERIAL_NO)", SerialNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.SHORT_DESC)", EquipShortName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EQ.EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.BRAND_ID", Brand, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.MODEL_ID", Model, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.EQUIP_SET_FLAG", EquipSet, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "SE.INSTALL_DATE", AppDateValue(InstallDataF), AppDateValue(InstallDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "EQ.EQUIPMENT_STATUS", EquipStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SE.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.BARCODE_NO)", BarCodeNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.SAP_MAT_CODE)", SAPNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.PART_NO)", PartNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.RD_NUMBER)", RDNumber.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.KIOSK_NO)", KioskNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.VERSION)", Version.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "EQ.QUANTITY", QuantityF, QuantityT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.UNIT_ID", Unit, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EQ.VENDOR_CODE)", EquipOwner.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "EQ.UNIT_COST", CostPerUnitF, CostPerUnitT, DBUTIL.FieldTypes.ftNumeric)

            If DAL.SQLDate(AppDateValue(WADateF)) & "" <> "NULL" AndAlso DAL.SQLDate(AppDateValue(WADateF)) & "" <> "NULL" Then
                If Criteria <> "" Then Criteria &= " AND "
                Criteria &= "((EQ.WA_DATE_START >= " & DAL.SQLDate(AppDateValue(WADateF)) & _
            " OR EQ.WA_DATE_START  <= " & DAL.SQLDate(AppDateValue(WADateF)) & ")" & _
            " AND EQ.WA_DATE_START <= " & DAL.SQLDate(AppDateValue(WADateT)) & " " & _
            " AND (EQ.WA_DATE_END >= " & DAL.SQLDate(AppDateValue(WADateT)) & " OR " & _
            "EQ.WA_DATE_END <= " & DAL.SQLDate(AppDateValue(WADateT)) & ")" & _
            " AND EQ.WA_DATE_END >=  " & DAL.SQLDate(AppDateValue(WADateF)) & ") "
            End If

            DB.AddCriteria(Criteria, "EQ.WARRANTY_TYPE", WAType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "EQ.PM_DATE", AppDateValue(PMDateF), AppDateValue(PMDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "EQ.UPDATE_PROG_DATE", AppDateValue(UpgradeDateF), AppDateValue(UpgradeDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "((SYSDATE-SE.INSTALL_DATE)/365)", UseAgeF, UseAgeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.NETWORK_ID", Network, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.SYSTEM_ID", System, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SE.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EQ.EQUIPMENT_SPEC)", Detail.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT EQ.*,ET.EQUIPMENT_TYPE_DESC,B.BRAND_NAME,M.MODEL_NAME,ES.EQUIPMENT_STATUS_DESC" & _
            ",SE.INSTALL_DATE AS SE_INSTALL_DATE" & _
            ",SE.NETWORK_ID,SE.LOCATION_ID,SE.SITE_ID,SE.SYSTEM_ID,S.SITE_NAME " & _
            " FROM EQUIPMENTS EQ,REF_EQUIPMENT_TYPES ET,REF_BRANDS B,REF_MODELS M " & _
            ",REF_EQUIPMENT_STATUS ES,SITE_EQUIPMENTS SE,SITES S,REF_PROVINCES P " & _
            "WHERE EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) AND EQ.BRAND_ID=B.BRAND_ID(+) AND EQ.MODEL_ID=M.MODEL_ID(+)" & _
            " AND EQ.EQUIPMENT_STATUS=ES.EQUIPMENT_STATUS(+) AND EQ.EQUIPMENT_ID=SE.EQUIPMENT_ID(+) AND SE.SITE_ID=S.SITE_ID(+) " & _
            " AND S.PROVINCE_ID=P.PROVINCE_ID(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EQ.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL, Conn, Trans)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEquipmentSet(Optional ByVal EquipSetID As String = "" _
    , Optional ByVal EquipID As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "ES.EQUIP_SET_ID", EquipSetID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "ES.EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)


            SQL = "SELECT ES.*,EQ.SERIAL_NO,EQ.EQUIPMENT_DESC,EQ.VENDOR_CODE,EQ.EQUIPMENT_STATUS,EQ.INSTALL_DATE,ET.EQUIPMENT_TYPE_DESC,B.BRAND_NAME,M.MODEL_NAME" & _
            ",RES.EQUIPMENT_STATUS_DESC,SE.SITE_ID,S.SITE_NAME,SE2.SITE_ID AS SITE_ID_EQ_SET " & _
            " FROM EQUIPMENT_SETS ES,EQUIPMENTS EQ,REF_EQUIPMENT_TYPES ET,REF_BRANDS B,REF_MODELS M " & _
            ",REF_EQUIPMENT_STATUS RES,SITE_EQUIPMENTS SE,SITES S,SITE_EQUIPMENTS SE2 " & _
            " WHERE ES.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) " & _
            " AND EQ.BRAND_ID=B.BRAND_ID(+) AND EQ.MODEL_ID=M.MODEL_ID(+)" & _
            " AND EQ.EQUIPMENT_STATUS=RES.EQUIPMENT_STATUS(+) AND EQ.EQUIPMENT_ID=SE.EQUIPMENT_ID(+) " & _
            " AND SE.SITE_ID=S.SITE_ID(+) AND ES.EQUIP_SET_ID=SE2.EQUIPMENT_ID(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY ES.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEquipMovement(Optional ByVal TranID As String = "" _
    , Optional ByVal TranDateF As String = "", Optional ByVal TranDateT As String = "" _
    , Optional ByVal EquipID As String = "", Optional ByVal ServiceID As String = "" _
    , Optional ByVal SiteID As String = "", Optional ByVal MovementType As String = "" _
    , Optional ByVal EquipName As String = "", Optional ByVal ServiceNo As String = "" _
    , Optional ByVal SiteName As String = "", Optional ByVal SerialNo As String = "" _
    , Optional ByVal ServiceType As String = "", Optional ByVal SiteRefID As String = "" _
    , Optional ByVal SiteRefName As String = "", Optional ByVal Remark As String = "" _
    , Optional ByVal RefTranID As String = "", Optional ByVal ServiceMmType As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "EM.TRANS_ID", TranID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "EM.TRANS_DATE", AppDateValue(TranDateF), AppDateValue(TranDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "EM.EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EM.SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EM.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EM.SITE_ID_OLD)", SiteRefID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EM.MOVEMENT_TYPE", MovementType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.SERIAL_NO)", SerialNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SV.SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EM.REF_TRANS_ID", RefTranID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SV.SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(S2.SITE_NAME)", SiteRefName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EM.COMMENTS)", Remark.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EM.SERVICE_MM_TYPE", ServiceMmType, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT EM.*,EQ.SERIAL_NO,EQ.EQUIPMENT_DESC,EQ.EQUIPMENT_STATUS,MT.MOVEMENT_TYPE_DESC,S.SITE_NAME,SV.SERVICE_NO,S2.SITE_NAME AS SITE_NAME_REF" & _
            ",VD.VENDOR_NAME,EQ.VENDOR_RESPONSE " & _
            " FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,REF_MOVEMENT_TYPES MT,SERVICES SV,SITES S,REF_PROVINCES P " & _
            " ,SITES S2,REF_PROVINCES P2,VENDORS VD" & _
            " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.MOVEMENT_TYPE=MT.MOVEMENT_TYPE(+) " & _
            " AND EM.SITE_ID=S.SITE_ID(+) AND EM.SERVICE_ID=SV.SERVICE_ID(+) " & _
            " AND S.PROVINCE_ID=P.PROVINCE_ID(+) AND EM.SITE_ID_OLD=S2.SITE_ID(+) AND S2.PROVINCE_ID=P2.PROVINCE_ID(+)" & _
            " AND EM.VENDOR_CODE=VD.VENDOR_CODE(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EM.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL, Conn, Trans)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageEquipment(ByVal op As Integer, Optional ByRef EquipID As String = Nothing, Optional ByVal EquipName As String = Nothing _
    , Optional ByVal EquipSName As String = Nothing, Optional ByVal SerialNo As String = Nothing _
    , Optional ByVal BarcodeNo As String = Nothing, Optional ByVal SAPNo As String = Nothing, Optional ByVal PartNo As String = Nothing _
    , Optional ByVal Brand As String = Nothing, Optional ByVal Model As String = Nothing, Optional ByVal Quantity As String = Nothing _
    , Optional ByVal Unit As String = Nothing, Optional ByVal CostUnit As String = Nothing, Optional ByVal WADateF As String = Nothing _
    , Optional ByVal WADateT As String = Nothing, Optional ByVal WAType As String = Nothing, Optional ByVal EquipOwner As String = Nothing _
    , Optional ByVal InstallDate As String = Nothing, Optional ByVal Longevity As String = Nothing _
    , Optional ByVal SiteID As String = Nothing, Optional ByVal LocationID As String = Nothing _
    , Optional ByVal NetworkID As String = Nothing, Optional ByVal IPAddr As String = Nothing _
    , Optional ByVal EquipStatus As String = Nothing, Optional ByVal EquipType As String = Nothing _
    , Optional ByVal Detail As String = Nothing, Optional ByVal EquipSetFlag As String = Nothing _
    , Optional ByVal PMDate As String = Nothing, Optional ByVal RDNumber As String = Nothing, Optional ByVal KioskNo As String = Nothing _
        , Optional ByVal Version As String = Nothing, Optional ByVal UpdatePGDate As String = Nothing _
        , Optional ByVal Seller As String = Nothing, Optional ByVal VendorResponse As String = Nothing, _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If EquipID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    EquipID = GenerateID("EQUIPMENTS", "EQUIPMENT_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                    'DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME"), DBUTIL.FieldTypes.ftText)
                    'DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_DESC", EquipName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SHORT_DESC", EquipSName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SERIAL_NO", SerialNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BARCODE_NO", BarcodeNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SAP_MAT_CODE", SAPNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PART_NO", PartNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "BRAND_ID", Brand, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "MODEL_ID", Model, DBUTIL.FieldTypes.ftNumeric)
                'DB.AddSQL2(op, SQL1, SQL2, "QUANTITY", Quantity, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "UNIT_ID", Unit, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "UNIT_COST", CostUnit, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "QUANTITY", Quantity, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "WA_DATE_START", AppDateValue(WADateF), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "WA_DATE_END", AppDateValue(WADateT), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "WARRANTY_TYPE", WAType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_CODE", EquipOwner, DBUTIL.FieldTypes.ftText)
                'DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
                'DB.AddSQL2(op, SQL1, SQL2, "LONGEVITY", Longevity, DBUTIL.FieldTypes.ftText)
                'DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                'DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
                'DB.AddSQL2(op, SQL1, SQL2, "NETWORK_ID", NetworkID, DBUTIL.FieldTypes.ftNumeric)
                'DB.AddSQL2(op, SQL1, SQL2, "IP_ADDRESS", IPAddr, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_STATUS", EquipStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_SPEC", Detail, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "EQUIP_SET_FLAG", EquipSetFlag, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PM_DATE", AppDateValue(PMDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "RD_NUMBER", RDNumber, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "KIOSK_NO", KioskNo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VERSION", Version, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "UPDATE_PROG_DATE", AppDateValue(UpdatePGDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SELLER", Seller, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_RESPONSE", VendorResponse, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", AppDateValue(InstallDate), DBUTIL.FieldTypes.ftDate)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "EQUIPMENTS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                EquipID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function ManageEquipmentSet(ByVal op As Integer, Optional ByVal EquipSetID As String = Nothing _
    , Optional ByVal EquipID As String = Nothing, Optional ByVal Quantity As String = Nothing, _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "EQUIP_SET_ID", EquipSetID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If op = DBUTIL.opINSERT Then
                    DB.AddSQL(op, SQL1, SQL2, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "EQUIP_SET_ID", EquipSetID, DBUTIL.FieldTypes.ftNumeric)
                    'DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME"), DBUTIL.FieldTypes.ftText)
                    'DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddCriteria(Criteria, "EQUIP_SET_ID", EquipSetID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "QUANTITY", Quantity, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "EQUIPMENT_SETS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                EquipID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function ManageEquipmentMovement(ByVal op As Integer, Optional ByRef TransID As String = Nothing _
    , Optional ByVal TransDateF As String = Nothing, Optional ByVal TransDateT As String = Nothing _
    , Optional ByVal MovementType As String = Nothing, Optional ByVal EquipID As String = Nothing _
    , Optional ByVal SiteID As String = Nothing, Optional ByVal OldSiteID As String = Nothing _
    , Optional ByVal Remark As String = Nothing, Optional ByVal ServiceID As String = Nothing _
    , Optional ByVal RefTransID As String = Nothing, Optional ByVal VendorCode As String = Nothing, _
    Optional ByVal LocationId As String = Nothing, Optional ByVal NetworkId As String = Nothing, _
    Optional ByVal InstallDate As String = Nothing, Optional ByVal SystemId As String = Nothing, _
    Optional ByVal ServiceMmType As String = Nothing, _
    Optional ByVal OtherCriteria As String = "", _
    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            Criteria = OtherCriteria
            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "SITE_ID_OLD", OldSiteID, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "REF_TRANS_ID", RefTransID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "TRANS_ID", TransID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If TransID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "TRANS_ID", TransID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    TransID = GenerateID("EQUIPMENT_MOVEMENTS", "TRANS_ID", Conn:=Conn, Trans:=Trans) & ""
                    DB.AddSQL(op, SQL1, SQL2, "TRANS_ID", TransID, DBUTIL.FieldTypes.ftNumeric)
                    'DB.AddSQL(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME"), DBUTIL.FieldTypes.ftText)
                    'DB.AddSQL(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "TRANS_DATE", AppDateValue(TransDateF), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "TRANS_DATE_TO", AppDateValue(TransDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "MOVEMENT_TYPE", MovementType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "EQUIPMENT_ID", EquipID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID_OLD", OldSiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "COMMENTS", Remark, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "REF_TRANS_ID", RefTransID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationId, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "NETWORK_ID", NetworkId, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "INSTALL_DATE", InstallDate, DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_ID", SystemId, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_MM_TYPE", ServiceMmType, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "EQUIPMENT_MOVEMENTS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                TransID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function SearchCheckStock(Optional ByVal CheckID As String = "" _
    , Optional ByVal CheckDateF As String = "", Optional ByVal CheckDateT As String = "" _
    , Optional ByVal SiteName As String = "", Optional ByVal VerifierName As String = "" _
    , Optional ByVal ProjectType As String = "" _
    , Optional ByVal Location As String = "", Optional ByVal SiteID As String = "" _
    , Optional ByVal Remark As String = "", Optional ByVal SerialNo As String = "" _
    , Optional ByVal EquipName As String = "", Optional ByVal EquipType As String = "" _
    , Optional ByVal CheckFound As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "CT.CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "CT.SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "CT.CHECK_DATE", AppDateValue(CheckDateF), AppDateValue(CheckDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(CT.VERIFIER_NAME)", VerifierName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "CT.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "CT.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(CT.REMARK)", Remark.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(EQ.SERIAL_NO)", SerialNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria2, "EQ.EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria2, "CSD.FOUND", CheckFound, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT CT.*,S.SITE_NAME,PT.PROJECT_TYPE_DESC,SL.LOCATION_NAME " & _
            " FROM CHECK_STOCKS CT,SITES S,REF_PROJECT_TYPES PT,SITE_LOCATIONS SL " & _
            " WHERE CT.SITE_ID=S.SITE_ID(+) AND CT.PROJECT_TYPE=PT.PROJECT_TYPE(+) " & _
            " AND CT.LOCATION_ID=SL.LOCATION_ID(+)"

            If Criteria2 <> "" Then SQL &= " AND CT.CHECK_ID IN (SELECT CSD.CHECK_ID FROM CHECK_STOCK_DETAILS CSD" & _
            ",EQUIPMENTS EQ WHERE CSD.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND " & Criteria2 & " )"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY CT.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchCheckStockDetail(Optional ByVal CheckID As String = "" _
    , Optional ByVal SiteID As String = "", Optional ByVal ProjectType As String = "" _
    , Optional ByVal Location As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(S.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "S.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SL.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)

            If CheckID <> "" Then
                'SQL = "SELECT SE.*,EQ.SERIAL_NO,EQ.EQUIPMENT_DESC,SL.LOCATION_NAME,CSD.EQUIPMENT_ID AS CHECK_EQ_ID " & _
                '" FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ,SITE_LOCATIONS SL,SITES S,SITE_SYSTEMS SS" & _
                '",(SELECT * FROM CHECK_STOCK_DETAILS WHERE CHECK_ID=" & CheckID & ") CSD " & _
                '" WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID AND SE.LOCATION_ID=SL.LOCATION_ID(+) " & _
                '"AND SE.SITE_ID=S.SITE_ID AND SE.EQUIPMENT_ID=CSD.EQUIPMENT_ID(+) AND S.SITE_ID=SS.SITE_ID(+) AND SE.SITE_ID=SL.SITE_ID(+) "
                SQL = "SELECT EQ.SERIAL_NO,EQ.EQUIPMENT_ID,EQ.EQUIPMENT_DESC,SL.LOCATION_NAME,CSD.FOUND " & _
                " FROM CHECK_STOCK_DETAILS CSD,CHECK_STOCKS CS,SITES S,EQUIPMENTS EQ, REF_EQUIPMENT_TYPES ET" & _
                ",SITE_LOCATIONS SL,SITE_EQUIPMENTS SE WHERE " & _
                " CSD.CHECK_ID=CS.CHECK_ID(+) AND CS.SITE_ID=S.SITE_ID AND CSD.EQUIPMENT_ID=EQ.EQUIPMENT_ID " & _
                "AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) " & _
                " AND CS.SITE_ID=SE.SITE_ID(+) AND CS.LOCATION_ID=SL.LOCATION_ID(+)"
            Else
                SQL = "SELECT SE.*,EQ.SERIAL_NO,EQ.EQUIPMENT_DESC,SL.LOCATION_NAME,0 AS FOUND" & _
                " FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ,SITE_LOCATIONS SL,SITES S " & _
                " WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID AND SE.LOCATION_ID=SL.LOCATION_ID(+) " & _
                "AND SE.SITE_ID=S.SITE_ID AND SE.SITE_ID=SL.SITE_ID(+)"
            End If

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If CheckID <> "" Then
                SQL &= " GROUP BY EQ.SERIAL_NO,EQ.EQUIPMENT_ID,EQ.EQUIPMENT_DESC,SL.LOCATION_NAME,CSD.FOUND"
            End If
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY EQ.SERIAL_NO"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchCheckStockSummary(Optional ByVal CheckID As String = "" _
    , Optional ByVal SiteID As String = "", Optional ByVal ProjectType As String = "" _
    , Optional ByVal Location As String = "", Optional ByVal SiteName As String = "" _
    , Optional ByVal CheckDateF As String = "", Optional ByVal CheckDateT As String = "" _
    , Optional ByVal EquipmentType As String = "", Optional ByVal Brand As String = "" _
    , Optional ByVal Model As String = "", Optional ByVal VerifierName As String = "" _
    , Optional ByVal Remark As String = "", Optional ByVal SerialNo As String = "" _
    , Optional ByVal EquipName As String = "", Optional ByVal CheckFound As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "" _
    , Optional ByVal StockBalance As Boolean = False) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "UPPER(S.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "S.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            If CheckID <> "" OrElse StockBalance Then
                DB.AddCriteria(Criteria, "CSD.CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "CS.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteriaRange(Criteria, "CS.CHECK_DATE", AppDateValue(CheckDateF), AppDateValue(CheckDateT), DBUTIL.FieldTypes.ftDate)
                DB.AddCriteria(Criteria, "CSD.FOUND", CheckFound, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "UPPER(CS.VERIFIER_NAME)", VerifierName.ToUpper, DBUTIL.FieldTypes.ftText)
                DB.AddCriteria(Criteria, "UPPER(CS.REMARK)", Remark.ToUpper, DBUTIL.FieldTypes.ftText)
            Else
                DB.AddCriteria(Criteria, "SE.LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
            End If
            DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "EQ.EQUIPMENT_TYPE", EquipmentType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.BRAND_ID", Brand, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "EQ.MODEL_ID", Model, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(EQ.SERIAL_NO)", SerialNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)

            If CheckID <> "" OrElse StockBalance Then
                'SQL = "SELECT EQ.EQUIPMENT_TYPE, ET.EQUIPMENT_TYPE_DESC, COUNT(SE.EQUIPMENT_ID) AS STOCK_CNT " & _
                '",COUNT(CSD.EQUIPMENT_ID) AS FOUND_CNT,COUNT(CSD.EQUIPMENT_ID)-COUNT(SE.EQUIPMENT_ID) AS DIFF_CNT " & _
                '" FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ, REF_EQUIPMENT_TYPES ET,SITE_LOCATIONS SL,SITES S " & _
                '",(SELECT CSD.*,CS.CHECK_DATE,CS.SITE_ID,CS.LOCATION_ID FROM CHECK_STOCK_DETAILS CSD,CHECK_STOCKS CS WHERE CS.CHECK_ID=CSD.CHECK_ID(+) AND " & _
                '"CS.CHECK_ID=" & CheckID & ") CSD " & _
                '" WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) " & _
                '"AND SE.LOCATION_ID=SL.LOCATION_ID(+) AND SE.SITE_ID=SL.SITE_ID(+) AND SE.SITE_ID=S.SITE_ID AND SE.EQUIPMENT_ID=CSD.EQUIPMENT_ID(+) "
                SQL = "SELECT EQ.EQUIPMENT_TYPE, ET.EQUIPMENT_TYPE_DESC, COUNT(CSD.EQUIPMENT_ID) AS STOCK_CNT " & _
                " ,SUM(CSD.FOUND) AS FOUND_CNT,SUM(CSD.FOUND)-COUNT(CSD.EQUIPMENT_ID) AS DIFF_CNT" & _
                " FROM CHECK_STOCK_DETAILS CSD,CHECK_STOCKS CS,SITES S,EQUIPMENTS EQ, REF_EQUIPMENT_TYPES ET" & _
                " WHERE " & _
                " CSD.CHECK_ID=CS.CHECK_ID(+) AND CS.SITE_ID=S.SITE_ID AND CSD.EQUIPMENT_ID=EQ.EQUIPMENT_ID " & _
                "AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+)"
            Else
                SQL = "SELECT EQ.EQUIPMENT_TYPE, ET.EQUIPMENT_TYPE_DESC, COUNT(SE.EQUIPMENT_ID) AS STOCK_CNT " & _
                ",0 AS FOUND_CNT,0-COUNT(SE.EQUIPMENT_ID) AS DIFF_CNT " & _
                " FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ, REF_EQUIPMENT_TYPES ET,SITE_LOCATIONS SL,SITES S " & _
                " WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) " & _
                "AND SE.LOCATION_ID=SL.LOCATION_ID(+) AND SE.SITE_ID=SL.SITE_ID(+) AND SE.SITE_ID=S.SITE_ID"
            End If

            'SQL = "SELECT EQ.EQUIPMENT_TYPE, ET.EQUIPMENT_TYPE_DESC, COUNT(SE.EQUIPMENT_ID) AS STOCK_CNT " & _
            '",COUNT(CSD.EQUIPMENT_ID) AS FOUND_CNT,COUNT(CSD.EQUIPMENT_ID)-COUNT(SE.EQUIPMENT_ID) AS DIFF_CNT " & _
            '" FROM SITE_EQUIPMENTS SE,EQUIPMENTS EQ, REF_EQUIPMENT_TYPES ET,SITE_LOCATIONS SL,SITES S " & _
            '",CHECK_STOCK_DETAILS CSD,CHECK_STOCKS CS " & _
            '" WHERE SE.EQUIPMENT_ID=EQ.EQUIPMENT_ID AND EQ.EQUIPMENT_TYPE=ET.EQUIPMENT_TYPE(+) " & _
            '"AND SE.LOCATION_ID=SL.LOCATION_ID(+) AND SE.SITE_ID=S.SITE_ID AND SE.EQUIPMENT_ID=CSD.EQUIPMENT_ID(+) " & _
            '"AND CSD.CHECK_ID=CS.CHECK_ID(+) AND SE.SITE_ID=SL.SITE_ID(+) "

            If Criteria <> "" Then SQL &= " AND " & Criteria
            SQL &= " GROUP BY EQ.EQUIPMENT_TYPE, ET.EQUIPMENT_TYPE_DESC"
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY ET.EQUIPMENT_TYPE_DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageCheckStock(ByVal op As Integer, Optional ByRef CheckID As String = Nothing _
    , Optional ByVal CheckDate As String = Nothing, Optional ByVal SiteID As String = Nothing _
    , Optional ByVal VerifierName As String = Nothing, Optional ByVal ProjectType As String = Nothing _
    , Optional ByVal Location As String = Nothing, Optional ByVal Remark As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If CheckID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    CheckID = GenerateID("CHECK_STOCKS", "CHECK_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "CHECK_DATE", AppDateValue(CheckDate), DBUTIL.FieldTypes.ftDate)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VERIFIER_NAME", VerifierName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", Location, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "REMARK", Remark, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "CHECK_STOCKS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                CheckID = ""
            End If
            Throw ex
        End Try

    End Function

    Public Function ManageCheckStockDetail(ByVal op As Integer, Optional ByRef CheckID As String = Nothing _
    , Optional ByVal EquipmentID As String = Nothing, Optional ByVal Found As String = Nothing _
    , Optional ByVal LocationID As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If op = DBUTIL.opUPDATE Then
                    DB.AddCriteria(Criteria, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    DB.AddSQL(op, SQL1, SQL2, "CHECK_ID", CheckID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "EQUIPMENT_ID", EquipmentID, DBUTIL.FieldTypes.ftNumeric)
                End If
            End If
            DB.AddSQL2(op, SQL1, SQL2, "FOUND", Found, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL2(op, SQL1, SQL2, "LOCATION_ID", LocationID, DBUTIL.FieldTypes.ftNumeric)
            SQL = DB.CombineSQL(op, SQL1, SQL2, "CHECK_STOCK_DETAILS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                CheckID = ""
            End If
            Throw ex
        End Try

    End Function
#End Region

#Region "Site"
#End Region

#Region "Service"
    'Public Function SearchServiceAction(Optional ByVal ServiceID As String = "", Optional ByVal ActionID As String = "", _
    'Optional ByVal ServiceNo As String = "", Optional ByVal ServiceType As String = "", Optional ByVal SiteID As String = "", _
    'Optional ByVal SiteName As String = "", Optional ByVal UserName As String = "", _
    'Optional ByVal UserGroup As String = "", Optional ByVal UserLevel As String = "", _
    'Optional ByVal ResponseDateF As String = "", Optional ByVal ResponseDateT As String = "", _
    'Optional ByVal ResolveDateF As String = "", Optional ByVal ResolveDateT As String = "", _
    'Optional ByVal CloseDateF As String = "", Optional ByVal CloseDateT As String = "", _
    'Optional ByVal Status As String = "", Optional ByVal ServiceStatusUpdate As String = "" _
    ', Optional ByVal AssignDateF As String = "", Optional ByVal AssignDateT As String = "", Optional ByVal AssignTo As String = "" _
    ', Optional ByVal SLAProfileA As String = "", Optional ByVal SeverityLevelA As String = "", Optional ByVal ResolutionTimeFA As String = "" _
    ', Optional ByVal ResolutionTimeTA As String = "", Optional ByVal ResponseTimeFA As String = "", Optional ByVal ResponseTimeTA As String = "" _
    ', Optional ByVal Note As String = "", Optional ByVal Reason As String = "", Optional ByVal RootCauseA As String = "" _
    ', Optional ByVal RequireDateF As String = "", Optional ByVal RequireDateT As String = "", Optional ByVal ProjectType As String = "" _
    ', Optional ByVal SiteGroup As String = "", Optional ByVal CallBy As String = "", Optional ByVal CallDetail As String = "" _
    ', Optional ByVal CallMethod As String = "", Optional ByVal CallMethodOther As String = "", Optional ByVal ProbCategory As String = "" _
    ', Optional ByVal ProbItem As String = "", Optional ByVal SLAProfile As String = "", Optional ByVal SeverityLevel As String = "" _
    ', Optional ByVal ServiceStatus As String = "", Optional ByVal ResolutionTimeF As String = "", Optional ByVal ResolutionTimeT As String = "" _
    ', Optional ByVal ResponseTimeF As String = "", Optional ByVal ResponseTimeT As String = "", Optional ByVal ClosingReason As String = "" _
    ', Optional ByVal RootCause As String = "", Optional ByVal UserGroupID As String = "" _
    ', Optional ByVal RowLimit As String = "", _
    'Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
    '    Dim DT As DataTable = Nothing
    '    Dim SQL As String = "", Criteria As String = ""
    '    Try
    '        Criteria = OtherCriteria
    '        DB.AddCriteria(Criteria, "SA.SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SA.ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(S.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria2Condi(Criteria, "UPPER(SA.USER_NAME)", "UPPER(SU.USER_DESC)", UserName.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SA.USER_GROUP_ID", UserGroupID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SU.GROUP_ID", UserGroup, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SU.USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "SA.RESPONSE_DATE", AppDateValue(ResponseDateF), AppDateValue(ResponseDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteriaRange(Criteria, "SA.RESOLVED_DATE", AppDateValue(ResolveDateF), AppDateValue(ResolveDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteriaRange(Criteria, "SV.CLOSE_DATE", AppDateValue(CloseDateF), AppDateValue(CloseDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria, "SA.SERVICE_STATUS", Status, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SA.SERVICE_STATUS_UPDATE", ServiceStatusUpdate, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "SA.ACTION_DATE", AppDateValue(AssignDateF), AppDateValue(AssignDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria2Condi(Criteria, "UPPER(SA.ASSIGN_TO)", "UPPER(SU2.USER_DESC)", AssignTo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SA.SLA_PROFILE_ID", SLAProfileA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SA.SEVERITY_LEVEL", SeverityLevelA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD.RESOLUTION_TIME/1440)", ResolutionTimeFA, ResolutionTimeTA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD.RESPONSE_TIME/1440)", ResponseTimeFA, ResponseTimeTA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SA.NOTE)", Note.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SA.REASON)", Reason.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SA.ROOT_CAUSE)", RootCauseA.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteriaRange(Criteria, "SV.REQUIRE_DATE", AppDateValue(RequireDateF), AppDateValue(RequireDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria, "SV.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SITE_GROUP_ID", SiteGroup, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.INFORMER_NAME)", CallBy.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CALL_DETAIL)", CallDetail.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.CALL_METHOD", CallMethod, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CALL_METHOD_OTHER)", CallMethodOther.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.PROBLEM_CATEGORY", ProbCategory, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.PROBLEM_ITEM", ProbItem, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SLA_PROFILE_ID", SLAProfile, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD3.RESOLUTION_TIME/1440)", ResolutionTimeF, ResolutionTimeT, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD3.RESPONSE_TIME/1440)", ResponseTimeF, ResponseTimeT, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CLOSE_REASON)", ClosingReason.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.ROOT_CAUSE)", RootCause.ToUpper, DBUTIL.FieldTypes.ftText)

    '        SQL = "SELECT SA.*,SS.SERVICE_STATUS_DESC,SD.RESPONSE_TIME,SD.RESPONSE_DAY,SD.RESPONSE_HOUR,SD.RESPONSE_MINUTE" & _
    '        ",SD.RESOLUTION_TIME,SD.RESOLUTION_DAY,SD.RESOLUTION_HOUR,SD.RESOLUTION_MINUTE,SU.ROLE_ID,SU.USER_TYPE,SU.USER_EMAIL,SU.GROUP_ID" & _
    '        ",UG.GROUP_NAME,DECODE(UG.GROUP_TYPE,NULL,UG.GROUP_TYPE,UG.GROUP_TYPE) AS GROUP_TYPE,UL.LEVEL_NAME,SV.SERVICE_NO,SV.SITE_ID,SV.CALL_DETAIL,SV.INFORMER_TEL " & _
    '        ",ST.SERVICE_TYPE_DESC,S.SITE_DESC,S.SITE_NAME,S.ADDRESS,RP.PROVINCE_NAME,S.ZIP_CODE,SV.CLOSE_DATE,SU.USER_TYPE,SV.CLOSE_DATE,SV.SITE_ID " & _
    '        ",DECODE(SU.USER_TYPE,1,'PTT Group','Vendor') AS USER_TYPE_DESC,SV.PROBLEM_CATEGORY,SV.PROBLEM_ITEM,SV.INFORMER_NAME" & _
    '        ",SV.SERVICE_DATE,PT.PROBLEM_TYPE_DESC,PR.PROB_RESOLVE_DESC " & _
    '        ",SD2.RESPONSE_TIME AS RESPONSE_TIME2,SD2.RESPONSE_DAY AS RESPONSE_DAY2,SD2.RESPONSE_HOUR AS RESPONSE_HOUR2" & _
    '        ",SD2.RESPONSE_MINUTE AS RESPONSE_MINUTE2" & _
    '        ",SD2.RESOLUTION_TIME AS RESOLUTION_TIME2,SD2.RESOLUTION_DAY AS RESOLUTION_DAY2,SD2.RESOLUTION_HOUR AS RESOLUTION_HOUR2" & _
    '        ",SD2.RESOLUTION_MINUTE AS RESOLUTION_MINUTE2,SS2.SERVICE_STATUS_DESC AS ACTION_SERVICE_STATUS_DESC" & _
    '        ",SD2.PROFILE_NAME,SD2.SEVERITY_LEVEL_DESC,UG2.GROUP_NAME AS ASSIGN_GROUP_NAME,UG2.GROUP_TYPE AS ASSIGN_GROUP_TYPE,SA2.RESPONSE_DATE AS ASSIGN_RESPONSE_DATE " & _
    '        ",SA3.REASON AS ASSIGN_REASON,SA3.ROOT_CAUSE AS ASSIGN_ROOT_CAUSE,SA3.NOTE AS ASSIGN_NOTE" & _
    '        ",PJT.PROJECT_TYPE_DESC,SAS.SLA AS OVER_SLA,DECODE(UG3.GROUP_NAME,NULL,UG.GROUP_NAME,UG3.GROUP_NAME) AS USER_GROUP_NAME" & _
    '        ",DECODE(UG4.GROUP_NAME,NULL,UG2.GROUP_NAME,UG4.GROUP_NAME ) AS ASSIGN_TO_GROUP_NAME" & _
    '        ",SU3.GROUP_ID AS ASSIGN_BY_GROUP_ID "
    '        'SQL += ",DECODE(SU.USER_TYPE,2,DECODE(SV.ISSUE_SLA,NULL,SD3.RESOLUTION_TIME-FLOOR((SYSDATE-SV.SERVICE_DATE)*24*60),SV.ISSUE_SLA)) AS SLA_CAL  " 
    '        SQL += "FROM SERVICE_ACTIONS SA,REF_SERVICE_STATUS SS,V_SLA_DETAILS SD,SYS_USERS SU,SYS_GROUPS UG,REF_USER_LEVELS UL " & _
    '        ",SERVICES SV,SITES S,REF_SERVICE_TYPES ST,REF_PROBLEM_TYPES PT,PROBLEM_RESOLVES PR,V_SLA_DETAILS SD2,SERVICE_ACTIONS SA2 " & _
    '        ",SYS_USERS SU2,SYS_GROUPS UG2,REF_SERVICE_STATUS SS2,REF_PROJECT_TYPES PJT,REF_PROVINCES RP,SERVICE_ACTIONS SA3 "
    '        'SQL += ",V_SLA_DETAILS SD3 "
    '        'SQL += "(SELECT SA.SERVICE_ID,SA.ACTION_ID FROM V_SERVICE_ACTION_SLAS SAS,SERVICE_ACTIONS SA WHERE " & _
    '        '"SA.SERVICE_ID=SAS.SERVICE_ID AND SAS.SLA < 0 GROUP BY SA.SERVICE_ID,SA.ACTION_ID) SAS "
    '        SQL += ",V_ALL_SV_ACTION_SLAS SAS,V_SLA_DETAILS SD3,SYS_GROUPS UG3,SYS_GROUPS UG4,SYS_USERS SU3 "
    '        SQL += "WHERE SA.SERVICE_STATUS=SS.SERVICE_STATUS(+) AND SA.ASSIGN_SLA_PROFILE_ID=SD.SLA_PROFILE_ID(+) " & _
    '        "AND SA.ASSIGN_SEVERITY_LEVEL=SD.SEVERITY_LEVEL(+) AND SA.USER_NAME=SU.USER_NAME(+) AND SU.GROUP_ID=UG.GROUP_ID(+) " & _
    '        " AND SU.USER_LEVEL=UL.LEVEL_ID(+) AND SA.SERVICE_ID=SV.SERVICE_ID AND SV.SITE_ID=S.SITE_ID(+) " & _
    '        "AND SV.SERVICE_TYPE=ST.SERVICE_TYPE(+) " & _
    '        "AND SV.PROBLEM_CATEGORY=PT.PROBLEM_TYPE(+) AND SV.PROBLEM_ITEM=PR.PROB_RESOLVE_ID(+) AND SA.SLA_PROFILE_ID=SD2.SLA_PROFILE_ID(+) " & _
    '        "AND SA.SEVERITY_LEVEL=SD2.SEVERITY_LEVEL(+) AND SA.ASSIGN_TO_ACTION_ID=SA2.ACTION_ID(+) AND SA.SERVICE_ID=SA2.SERVICE_ID(+) " & _
    '        "AND SA.ASSIGN_TO=SU2.USER_NAME(+) AND SU2.GROUP_ID=UG2.GROUP_ID(+) AND SA.SERVICE_STATUS_UPDATE=SS2.SERVICE_STATUS(+) " & _
    '        "AND SV.PROJECT_TYPE=PJT.PROJECT_TYPE(+) AND S.PROVINCE_ID=RP.PROVINCE_ID(+) AND SA.SERVICE_ID=SA3.SERVICE_ID(+) AND SA.REF_ACTION_ID=SA3.ACTION_ID(+) "
    '        'SQL += " AND SV.SLA_PROFILE_ID=SD3.SLA_PROFILE_ID(+) AND SV.SEVERITY_LEVEL=SD3.SEVERITY_LEVEL(+)"
    '        SQL += " AND SA.SERVICE_ID=SAS.SERVICE_ID(+) AND SA.ACTION_ID=SAS.ACTION_ID(+) AND SV.SLA_PROFILE_ID=SD3.SLA_PROFILE_ID(+) " & _
    '        "AND SV.SEVERITY_LEVEL=SD3.SEVERITY_LEVEL(+)  AND SA.USER_GROUP_ID=UG3.GROUP_ID(+) AND SA.ASSIGN_TO_GRP=UG4.GROUP_ID(+) " & _
    '        "AND SA.ASSIGN_BY=SU3.USER_NAME(+) "

    '        If Criteria <> "" Then SQL &= " AND " & Criteria
    '        If RowLimit <> "" Then SQL &= " AND ROWNUM <= " & RowLimit
    '        If OrderBy <> "" Then
    '            SQL &= " ORDER BY " & OrderBy
    '        Else
    '            SQL &= " ORDER BY SA.ACTION_DATE,SA.ACTION_ID"
    '        End If
    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Public Function SearchServiceAction(Optional ByVal ServiceID As String = "", Optional ByVal ActionID As String = "", _
    Optional ByVal ServiceNo As String = "", Optional ByVal ServiceType As String = "", Optional ByVal SiteID As String = "", _
    Optional ByVal SiteName As String = "", Optional ByVal UserName As String = "", _
    Optional ByVal UserGroup As String = "", Optional ByVal UserLevel As String = "", _
    Optional ByVal ResponseDateF As String = "", Optional ByVal ResponseDateT As String = "", _
    Optional ByVal ResolveDateF As String = "", Optional ByVal ResolveDateT As String = "", _
    Optional ByVal CloseDateF As String = "", Optional ByVal CloseDateT As String = "", _
    Optional ByVal Status As String = "", Optional ByVal ServiceStatusUpdate As String = "" _
    , Optional ByVal AssignDateF As String = "", Optional ByVal AssignDateT As String = "", Optional ByVal AssignTo As String = "" _
    , Optional ByVal SLAProfileA As String = "", Optional ByVal SeverityLevelA As String = "", Optional ByVal ResolutionTimeFA As String = "" _
    , Optional ByVal ResolutionTimeTA As String = "", Optional ByVal ResponseTimeFA As String = "", Optional ByVal ResponseTimeTA As String = "" _
    , Optional ByVal Note As String = "", Optional ByVal Reason As String = "", Optional ByVal RootCauseA As String = "" _
    , Optional ByVal RequireDateF As String = "", Optional ByVal RequireDateT As String = "", Optional ByVal ProjectType As String = "" _
    , Optional ByVal SiteGroup As String = "", Optional ByVal CallBy As String = "", Optional ByVal CallDetail As String = "" _
    , Optional ByVal CallMethod As String = "", Optional ByVal CallMethodOther As String = "", Optional ByVal ProbCategory As String = "" _
    , Optional ByVal ProbItem As String = "", Optional ByVal SLAProfile As String = "", Optional ByVal SeverityLevel As String = "" _
    , Optional ByVal ServiceStatus As String = "", Optional ByVal ResolutionTimeF As String = "", Optional ByVal ResolutionTimeT As String = "" _
    , Optional ByVal ResponseTimeF As String = "", Optional ByVal ResponseTimeT As String = "", Optional ByVal ClosingReason As String = "" _
    , Optional ByVal RootCause As String = "", Optional ByVal UserGroupID As String = "" _
    , Optional ByVal RowLimit As String = "", _
    Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""
        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria, "UPPER(USER_NAME)", "UPPER(USER_DESC)", UserName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "USER_GROUP_ID", UserGroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "GROUP_ID", UserGroup, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "USER_LEVEL", UserLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "RESPONSE_DATE", AppDateValue(ResponseDateF), AppDateValue(ResponseDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "RESOLVED_DATE", AppDateValue(ResolveDateF), AppDateValue(ResolveDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "CLOSE_DATE", AppDateValue(CloseDateF), AppDateValue(CloseDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "SERVICE_STATUS", Status, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SERVICE_STATUS_UPDATE", ServiceStatusUpdate, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "ACTION_DATE", AppDateValue(AssignDateF), AppDateValue(AssignDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria2Condi(Criteria, "UPPER(ASSIGN_TO)", "UPPER(SU2.USER_DESC)", AssignTo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfileA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevelA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "(RESOLUTION_TIME/1440)", ResolutionTimeFA, ResolutionTimeTA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "(RESPONSE_TIME/1440)", ResponseTimeFA, ResponseTimeTA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(NOTE)", Note.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(REASON)", Reason.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(ROOT_CAUSE)", RootCauseA.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "REQUIRE_DATE", AppDateValue(RequireDateF), AppDateValue(RequireDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SITE_GROUP_ID", SiteGroup, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(INFORMER_NAME)", CallBy.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(CALL_DETAIL)", CallDetail.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "CALL_METHOD", CallMethod, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(CALL_METHOD_OTHER)", CallMethodOther.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PROBLEM_CATEGORY", ProbCategory, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "PROBLEM_ITEM", ProbItem, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SV_SLA_PROFILE_ID", SLAProfile, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SV_SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SV_SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "SA_RESOLUTION_TIME", ResolutionTimeF, ResolutionTimeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "SA_RESPONSE_TIME", ResponseTimeF, ResponseTimeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(CLOSE_REASON)", ClosingReason.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SV_ROOT_CAUSE)", RootCause.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT * FROM V_SERVICE_ACTIONS "

            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If RowLimit <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= "ROWNUM <= " & RowLimit
            End If
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY ACTION_DATE,ACTION_ID ASC"
            End If
            DB.OpenDT(DT, SQL)
            'Me.SqlLog(SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchServiceOverSLA(Optional ByVal RowLimit As String = "", _
    Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""
        Try
            Criteria = OtherCriteria
            'DB.AddCriteria(Criteria, "SA.SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteriaRange(Criteria, "SA.RESPONSE_DATE", AppDateValue(ResponseDateF), AppDateValue(ResponseDateT), DBUTIL.FieldTypes.ftDateTime)

            SQL = "SELECT SAS.*,SV.SITE_ID FROM V_SERVICE_ACTION_SLAS SAS,SERVICES SV WHERE SAS.SERVICE_ID=SV.SERVICE_ID(+)"
            If Criteria <> "" Then SQL &= " AND " & Criteria
            If RowLimit <> "" Then
                'If Criteria <> "" Then
                '    SQL &= " AND "
                'Else
                '    SQL &= " WHERE "
                'End If
                If Criteria <> "" Then
                    SQL &= " AND "
                End If
                SQL &= " ROWNUM <= " & RowLimit
            End If
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY SAS.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchServiceAttachment(Optional ByVal ServiceID As String = "", Optional ByVal AttachmentID As String = "", _
    Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "ATTACHMENT_ID", AttachmentID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * " & _
            "FROM SERVICE_ATTACHMENTS "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Public Function SearchService(Optional ByVal ServiceID As String = "", Optional ByVal ServiceNo As String = "" _
    '    , Optional ByVal ServiceDateF As String = "", Optional ByVal ServiceDateT As String = "", Optional ByVal ProjectType As String = "" _
    '    , Optional ByVal SiteGroupID As String = "", Optional ByVal SiteName As String = "" _
    '    , Optional ByVal CallBy As String = "", Optional ByVal CallDetail As String = "" _
    '    , Optional ByVal Category As String = "", Optional ByVal SLAProfile As String = "" _
    '    , Optional ByVal SeverityLevel As String = "", Optional ByVal ServiceStatus As String = "" _
    '    , Optional ByVal SiteID As String = "", Optional ByVal ServiceType As String = "", Optional ByVal TelNo As String = "" _
    '    , Optional ByVal Email As String = "", Optional ByVal RequireDateF As String = "", Optional ByVal RequireDateT As String = "" _
    '    , Optional ByVal CallMethod As String = "", Optional ByVal CallMethodOther As String = "", Optional ByVal ProbItem As String = "" _
    '    , Optional ByVal ResolutionTimeF As String = "", Optional ByVal ResolutionTimeT As String = "" _
    '    , Optional ByVal ResponseTimeF As String = "", Optional ByVal ResponseTimeT As String = "" _
    '    , Optional ByVal ClosingReason As String = "", Optional ByVal RootCause As String = "", Optional ByVal Description As String = "" _
    '    , Optional ByVal AssignBy As String = "", Optional ByVal AssignDateF As String = "", Optional ByVal AssignDateT As String = "" _
    '    , Optional ByVal ResponseDateF As String = "", Optional ByVal ResponseDateT As String = "", Optional ByVal AssignStatus As String = "" _
    '    , Optional ByVal ResolveDateF As String = "", Optional ByVal ResolveDateT As String = "" _
    '    , Optional ByVal CloseDateF As String = "", Optional ByVal CloseDateT As String = "", Optional ByVal AssignTo As String = "" _
    '    , Optional ByVal SLAProfileA As String = "", Optional ByVal SeverityLevelA As String = "" _
    '    , Optional ByVal ResolutionTimeFA As String = "", Optional ByVal ResolutionTimeTA As String = "" _
    '    , Optional ByVal ResponseTimeFA As String = "", Optional ByVal ResponseTimeTA As String = "" _
    '    , Optional ByVal Note As String = "", Optional ByVal Reason As String = "", Optional ByVal RootCauseA As String = "" _
    '    , Optional ByVal SerailNo As String = "", Optional ByVal EquipName As String = "", Optional ByVal SAPNo As String = "" _
    '    , Optional ByVal EquipType As String = "", Optional ByVal SiteIDEI As String = "", Optional ByVal SiteNameEI As String = "" _
    '    , Optional ByVal SerailNoER As String = "", Optional ByVal EquipNameER As String = "", Optional ByVal SAPNoER As String = "" _
    '    , Optional ByVal EquipTypeER As String = "", Optional ByVal SiteIDER As String = "", Optional ByVal SiteNameER As String = "" _
    '    , Optional ByVal RowLimit As String = "", Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
    '    Dim DT As DataTable = Nothing
    '    Dim SQL As String = "", Criteria As String = "", Criteria2 As String = "", Criteria3 As String = "" _
    '    , Criteria4 As String = "", Criteria5 As String = ""

    '    Try
    '        Criteria = OtherCriteria
    '        DB.AddCriteria(Criteria, "SV.SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteriaRange(Criteria, "SV.SERVICE_DATE", AppDateValue(ServiceDateF), AppDateValue(ServiceDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria, "SV.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(S.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.INFORMER_NAME)", CallBy.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CALL_DETAIL)", CallDetail.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.PROBLEM_CATEGORY", Category, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SLA_PROFILE_ID", SLAProfile, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "SV.SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.INFORMER_TEL)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.INFORMER_EMAIL)", Email.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteriaRange(Criteria, "SV.REQUIRE_DATE", AppDateValue(RequireDateF), AppDateValue(RequireDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria, "SV.CALL_METHOD", CallMethod, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CALL_METHOD_OTHER)", CallMethodOther.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "SV.PROBLEM_ITEM", ProbItem, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD.RESOLUTION_TIME/1440)", ResolutionTimeF, ResolutionTimeT, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria, "(SD.RESPONSE_TIME/1440)", ResponseTimeF, ResponseTimeT, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria, "UPPER(SV.CLOSE_REASON)", ClosingReason.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria, "UPPER(SV.ROOT_CAUSE)", RootCause.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteriaRange(Criteria2, "SV.CLOSE_DATE", AppDateValue(CloseDateF), AppDateValue(CloseDateT), DBUTIL.FieldTypes.ftDate)

    '        DB.AddCriteria(Criteria2, "UPPER(ATTACHMENT_DESC)", Description.ToUpper, DBUTIL.FieldTypes.ftText)

    '        DB.AddCriteria2Condi(Criteria3, "UPPER(SA.ASSIGN_BY)", "UPPER(SU.USER_DESC)", AssignBy.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteriaRange(Criteria3, "SA.ACTION_DATE", AppDateValue(AssignDateF), AppDateValue(AssignDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteriaRange(Criteria3, "SA.RESPONSE_DATE", AppDateValue(ResponseDateF), AppDateValue(ResponseDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria(Criteria3, "SA.SERVICE_STATUS", AssignStatus, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria3, "SA.RESOLVED_DATE", AppDateValue(ResolveDateF), AppDateValue(ResolveDateT), DBUTIL.FieldTypes.ftDate)
    '        DB.AddCriteria2Condi(Criteria3, "UPPER(SA.USER_NAME)", "UPPER(SU2.USER_DESC)", AssignTo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria3, "SA.SLA_PROFILE_ID", SLAProfileA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria3, "SA.SEVERITY_LEVEL", SeverityLevelA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria3, "(SD.RESOLUTION_TIME/1440)", ResolutionTimeFA, ResolutionTimeTA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(Criteria3, "(SD.RESPONSE_TIME/1440)", ResponseTimeFA, ResponseTimeTA, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria3, "UPPER(SA.NOTE)", Note.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria3, "UPPER(SA.REASON)", Reason.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria3, "UPPER(SA.ROOT_CAUSE)", RootCauseA.ToUpper, DBUTIL.FieldTypes.ftText)

    '        DB.AddCriteria(Criteria4, "UPPER(EQ.SERIAL_NO)", SerailNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria2Condi(Criteria4, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria4, "UPPER(EQ.SAP_MAT_CODE)", SAPNo.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria4, "EQ.EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria4, "UPPER(S.SITE_ID)", SiteIDEI.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria2Condi(Criteria4, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME)", SiteNameEI.ToUpper, DBUTIL.FieldTypes.ftText)

    '        DB.AddCriteria(Criteria5, "UPPER(EQ.SERIAL_NO)", SerailNoER.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria2Condi(Criteria5, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipNameER.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria5, "UPPER(EQ.SAP_MAT_CODE)", SAPNoER.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria(Criteria5, "EQ.EQUIPMENT_TYPE", EquipTypeER, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteria(Criteria5, "UPPER(S.SITE_ID)", SiteIDER.ToUpper, DBUTIL.FieldTypes.ftText)
    '        DB.AddCriteria2Condi(Criteria5, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME)", SiteNameER.ToUpper, DBUTIL.FieldTypes.ftText)

    '        SQL = "SELECT SV.*,SS.SERVICE_STATUS_DESC,PT.PROJECT_TYPE_DESC,S.SITE_NAME,SVT.SERVICE_TYPE_DESC,SD.RESOLUTION_TIME" & _
    '        ",SD.RESPONSE_TIME,SG.SITE_GROUP_NAME,SD.PROFILE_NAME2,SD.SEVERITY_LEVEL_DESC,PB.PROBLEM_TYPE_DESC,PBR.PROB_RESOLVE_DESC " & _
    '        "FROM SERVICES SV,REF_SERVICE_STATUS SS,REF_PROJECT_TYPES PT" & _
    '        ",SITES S,REF_SERVICE_TYPES SVT,V_SLA_DETAILS SD,SITE_GROUPS SG,REF_PROBLEM_TYPES PB,PROBLEM_RESOLVES PBR " & _
    '        " WHERE SV.SERVICE_STATUS=SS.SERVICE_STATUS(+) AND SV.PROJECT_TYPE=PT.PROJECT_TYPE(+) AND SV.SITE_ID=S.SITE_ID(+)" & _
    '        " AND SV.SERVICE_TYPE=SVT.SERVICE_TYPE(+) AND SV.SLA_PROFILE_ID=SD.SLA_PROFILE_ID(+) " & _
    '        "AND SV.SEVERITY_LEVEL=SD.SEVERITY_LEVEL(+) AND SV.SITE_GROUP_ID=SG.SITE_GROUP_ID(+)" & _
    '        "AND SV.PROBLEM_CATEGORY=PB.PROBLEM_TYPE(+) AND SV.PROBLEM_ITEM=PBR.PROB_RESOLVE_ID(+)"
    '        If Criteria <> "" Then SQL &= " AND " & Criteria

    '        If Criteria2 <> "" Then SQL &= " AND SV.SERVICE_ID IN (SELECT SERVICE_ID FROM SERVICE_ATTACHMENTS WHERE " & Criteria2 & ")"

    '        If Criteria3 <> "" Then SQL &= " AND SV.SERVICE_ID IN (SELECT SA.SERVICE_ID FROM SERVICE_ACTIONS SA,SYS_USERS SU,SYS_USERS SU2 " & _
    '        ",V_SLA_DETAILS SD WHERE SA.ASSIGN_BY=SU.USER_NAME(+) AND SA.USER_NAME=SU2.USER_NAME(+) " & _
    '        "AND SA.SLA_PROFILE_ID=SD.SLA_PROFILE_ID(+) AND SA.SEVERITY_LEVEL=SD.SEVERITY_LEVEL(+) AND " & Criteria3 & ")"

    '        If Criteria4 <> "" Then SQL &= " AND SV.SERVICE_ID IN (SELECT EM.SERVICE_ID FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,SITES S" & _
    '        " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID_OLD=S.SITE_ID(+) AND EM.MOVEMENT_TYPE=1" & _
    '        " AND " & Criteria4 & ")"

    '        If Criteria5 <> "" Then SQL &= " AND SV.SERVICE_ID IN (SELECT EM.SERVICE_ID FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,SITES S" & _
    '        " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID=S.SITE_ID(+) AND EM.MOVEMENT_TYPE=3" & _
    '        " AND " & Criteria5 & ")"

    '        If RowLimit <> "" Then SQL &= " AND ROWNUM <= " & RowLimit
    '        If OrderBy <> "" Then
    '            SQL &= " ORDER BY " & OrderBy
    '        Else
    '            SQL &= " ORDER BY SV.DATE_UPDATED DESC"
    '        End If
    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Public Function SearchService(Optional ByVal ServiceID As String = "", Optional ByVal ServiceNo As String = "" _
        , Optional ByVal ServiceDateF As String = "", Optional ByVal ServiceDateT As String = "", Optional ByVal ProjectType As String = "" _
        , Optional ByVal SiteGroupID As String = "", Optional ByVal SiteName As String = "" _
        , Optional ByVal CallBy As String = "", Optional ByVal CallDetail As String = "" _
        , Optional ByVal Category As String = "", Optional ByVal SLAProfile As String = "" _
        , Optional ByVal SeverityLevel As String = "", Optional ByVal ServiceStatus As String = "" _
        , Optional ByVal SiteID As String = "", Optional ByVal ServiceType As String = "", Optional ByVal TelNo As String = "" _
        , Optional ByVal Email As String = "", Optional ByVal RequireDateF As String = "", Optional ByVal RequireDateT As String = "" _
        , Optional ByVal CallMethod As String = "", Optional ByVal CallMethodOther As String = "", Optional ByVal ProbItem As String = "" _
        , Optional ByVal ResolutionTimeF As String = "", Optional ByVal ResolutionTimeT As String = "" _
        , Optional ByVal ResponseTimeF As String = "", Optional ByVal ResponseTimeT As String = "" _
        , Optional ByVal ClosingReason As String = "", Optional ByVal RootCause As String = "", Optional ByVal Description As String = "" _
        , Optional ByVal AssignBy As String = "", Optional ByVal AssignDateF As String = "", Optional ByVal AssignDateT As String = "" _
        , Optional ByVal ResponseDateF As String = "", Optional ByVal ResponseDateT As String = "", Optional ByVal AssignStatus As String = "" _
        , Optional ByVal ResolveDateF As String = "", Optional ByVal ResolveDateT As String = "" _
        , Optional ByVal CloseDateF As String = "", Optional ByVal CloseDateT As String = "", Optional ByVal AssignTo As String = "" _
        , Optional ByVal SLAProfileA As String = "", Optional ByVal SeverityLevelA As String = "" _
        , Optional ByVal ResolutionTimeFA As String = "", Optional ByVal ResolutionTimeTA As String = "" _
        , Optional ByVal ResponseTimeFA As String = "", Optional ByVal ResponseTimeTA As String = "" _
        , Optional ByVal Note As String = "", Optional ByVal Reason As String = "", Optional ByVal RootCauseA As String = "" _
        , Optional ByVal SerailNo As String = "", Optional ByVal EquipName As String = "", Optional ByVal SAPNo As String = "" _
        , Optional ByVal EquipType As String = "", Optional ByVal SiteIDEI As String = "", Optional ByVal SiteNameEI As String = "" _
        , Optional ByVal SerailNoER As String = "", Optional ByVal EquipNameER As String = "", Optional ByVal SAPNoER As String = "" _
        , Optional ByVal EquipTypeER As String = "", Optional ByVal SiteIDER As String = "", Optional ByVal SiteNameER As String = "" _
        , Optional ByVal RowLimit As String = "", Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "", Optional ByVal Cri As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = "", Criteria3 As String = "" _
        , Criteria4 As String = "", Criteria5 As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "SERVICE_DATE", AppDateValue(ServiceDateF), AppDateValue(ServiceDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(INFORMER_NAME)", CallBy.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(CALL_DETAIL)", CallDetail.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PROBLEM_CATEGORY", Category, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SLA_PROFILE_ID", SLAProfile, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(V_SERVICES.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(INFORMER_TEL)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(INFORMER_EMAIL)", Email.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "REQUIRE_DATE", AppDateValue(RequireDateF), AppDateValue(RequireDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "CALL_METHOD", CallMethod, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(CALL_METHOD_OTHER)", CallMethodOther.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "PROBLEM_ITEM", ProbItem, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "(RESOLUTION_TIME/1440)", ResolutionTimeF, ResolutionTimeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "(RESPONSE_TIME/1440)", ResponseTimeF, ResponseTimeT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(CLOSE_REASON)", ClosingReason.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(ROOT_CAUSE)", RootCause.ToUpper, DBUTIL.FieldTypes.ftText)

            DB.AddCriteriaRange(Criteria2, "SV.CLOSE_DATE", AppDateValue(CloseDateF), AppDateValue(CloseDateT), DBUTIL.FieldTypes.ftDate)

            DB.AddCriteria(Criteria2, "UPPER(ATTACHMENT_DESC)", Description.ToUpper, DBUTIL.FieldTypes.ftText)

            DB.AddCriteria2Condi(Criteria3, "UPPER(SA.ASSIGN_BY)", "UPPER(SU.USER_DESC)", AssignBy.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria3, "SA.ACTION_DATE", AppDateValue(AssignDateF), AppDateValue(AssignDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria3, "SA.RESPONSE_DATE", AppDateValue(ResponseDateF), AppDateValue(ResponseDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria3, "SA.SERVICE_STATUS", AssignStatus, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria3, "SA.RESOLVED_DATE", AppDateValue(ResolveDateF), AppDateValue(ResolveDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria2Condi(Criteria3, "UPPER(SA.USER_NAME)", "UPPER(SU2.USER_DESC)", AssignTo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria3, "SA.SLA_PROFILE_ID", SLAProfileA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria3, "SA.SEVERITY_LEVEL", SeverityLevelA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria3, "(SD.RESOLUTION_TIME/1440)", ResolutionTimeFA, ResolutionTimeTA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria3, "(SD.RESPONSE_TIME/1440)", ResponseTimeFA, ResponseTimeTA, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria3, "UPPER(SA.NOTE)", Note.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria3, "UPPER(SA.REASON)", Reason.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria3, "UPPER(SA.ROOT_CAUSE)", RootCauseA.ToUpper, DBUTIL.FieldTypes.ftText)

            DB.AddCriteria(Criteria4, "UPPER(EQ.SERIAL_NO)", SerailNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria4, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria4, "UPPER(EQ.SAP_MAT_CODE)", SAPNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria4, "EQ.EQUIPMENT_TYPE", EquipType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria4, "UPPER(S.SITE_ID)", SiteIDEI.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria4, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME)", SiteNameEI.ToUpper, DBUTIL.FieldTypes.ftText)

            DB.AddCriteria(Criteria5, "UPPER(EQ.SERIAL_NO)", SerailNoER.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria5, "UPPER(EQ.SHORT_DESC)", "UPPER(EQ.EQUIPMENT_DESC)", EquipNameER.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria5, "UPPER(EQ.SAP_MAT_CODE)", SAPNoER.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria5, "EQ.EQUIPMENT_TYPE", EquipTypeER, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria5, "UPPER(S.SITE_ID)", SiteIDER.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria2Condi(Criteria5, "UPPER(S.SITE_NAME)", "UPPER(S.SITE_NAME)", SiteNameER.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = " SELECT " & _
                " V_SERVICES.SERVICE_ID, V_SERVICES.SERVICE_NO, V_SERVICES.SERVICE_DATE, V_SERVICES.SITE_ID, V_SERVICES.PROJECT_TYPE, V_SERVICES.SITE_TYPE, " & _
                " V_SERVICES.INFORMER_TITLE_ID, V_SERVICES.INFORMER_NAME, V_SERVICES.INFORMER_TEL, V_SERVICES.INFORM_DATE, V_SERVICES.INFORMER_EMAIL,  " & _
                " V_SERVICES.REF_CALL_NUMBER, V_SERVICES.CALL_DETAIL, V_SERVICES.PROBLEM_CATEGORY, V_SERVICES.PROBLEM_ITEM, V_SERVICES.SLA_PROFILE_ID,  " & _
                " V_SERVICES.SEVERITY_LEVEL, V_SERVICES.SERVICE_STATUS, V_SERVICES.CLOSE_DATE, V_SERVICES.DATE_CREATED, V_SERVICES.DATE_UPDATED,  " & _
                " V_SERVICES.SERVICE_TYPE, V_SERVICES.VENDOR_SLA, V_SERVICES.ISSUE_SLA, V_SERVICES.TOTAL_SERVICE_TIME, V_SERVICES.FVENDOR_SLA, " & _
                " V_SERVICES.FISSUE_SLA, V_SERVICES.FTOTAL_SERVICE_TIME, V_SERVICES.REQUIRE_DATE, V_SERVICES.CALL_METHOD_OTHER, " & _
                " V_SERVICES.SITE_GROUP_ID, V_SERVICES.CALL_METHOD, V_SERVICES.ROOT_CAUSE, V_SERVICES.CLOSE_REASON, V_SERVICES.PENDING_REASON, " & _
                " V_SERVICES.REJECT_REASON, V_SERVICES.RESOLVED_REASON, V_SERVICES.SERVICE_STATUS_DESC, V_SERVICES.PROJECT_TYPE_DESC,  " & _
                " V_SERVICES.SITE_NAME, V_SERVICES.TEL_NO, V_SERVICES.FAX_NO, V_SERVICES.ADDRESS, V_SERVICES.ZIP_CODE, V_SERVICES.PROVINCE_ID, " & _
                " V_SERVICES.PROVINCE_NAME, V_SERVICES.SERVICE_TYPE_DESC, V_SERVICES.RESOLUTION_TIME, V_SERVICES.RESPONSE_TIME,  " & _
                " V_SERVICES.PROFILE_NAME2, V_SERVICES.SEVERITY_LEVEL_DESC, V_SERVICES.SITE_GROUP_NAME, V_SERVICES.PROBLEM_TYPE, " & _
                " V_SERVICES.PROBLEM_TYPE_DESC, V_SERVICES.TOTAL_PERCENT, V_SERVICES.PROB_RESOLVE_DESC, V_SERVICES.USER_CREATED,  " & _
                " V_SERVICES.VENDOR_ACTION, V_SERVICES.GROUP_VENDOR_ACTION, V_SERVICES.RESOLVED_DATE, V_SERVICES.SALE_AREA,  " & _
                " V_SERVICES.SALE_AREA_NAME, V_SERVICES.FALL_SERVICE_TIME, CMP.COMPANY, SLASTATUS2(V_SERVICES.SERVICE_ID, V_SERVICES.RESOLUTION_TIME) AS FSLA_STATUS " & _
                ", cmp.COMPANY, " & _
                " SLASTATUS2( " & _
                " V_SERVICES.SERVICE_ID, V_SERVICES.RESOLUTION_TIME " & _
                " ) AS fSLA_STATUS " & _
                " FROM     V_SERVICES LEFT OUTER JOIN" & _
                " (SELECT  SITE_ID, OWNER_TYPE_DESC AS COMPANY" & _
                " FROM   (SELECT  SS.SITE_ID, SS.SYSTEM_NAME, WT.OWNER_TYPE_DESC, RSS.SYSTEM_STATUS_DESC" & _
                " FROM   SITE_SYSTEMS SS INNER JOIN" & _
                "    REF_OWNER_TYPES WT ON SS.OWNER_TYPE = WT.OWNER_TYPE INNER JOIN" & _
                "    REF_SYSTEM_STATUS RSS ON SS.SYSTEM_STATUS = RSS.SYSTEM_STATUS" & _
                " WHERE        (RSS.SYSTEM_STATUS = 1)) derivedtbl_1) cmp ON cmp.SITE_ID = V_SERVICES.SITE_ID "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria

            If Criteria2 <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= " SERVICE_ID IN (SELECT SERVICE_ID FROM SERVICE_ATTACHMENTS WHERE " & Criteria2 & ")"
            End If

            If Criteria3 <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= " SERVICE_ID IN (SELECT SA.SERVICE_ID FROM SERVICE_ACTIONS SA,SYS_USERS SU,SYS_USERS SU2 " & _
            ",V_SLA_DETAILS SD WHERE SA.ASSIGN_BY=SU.USER_NAME(+) AND SA.USER_NAME=SU2.USER_NAME(+) " & _
            "AND SA.SLA_PROFILE_ID=SD.SLA_PROFILE_ID(+) AND SA.SEVERITY_LEVEL=SD.SEVERITY_LEVEL(+) AND " & Criteria3 & ")"
            End If

            If Criteria4 <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= " SERVICE_ID IN (SELECT EM.SERVICE_ID FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,SITES S" & _
            " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID_OLD=S.SITE_ID(+) AND EM.MOVEMENT_TYPE=1" & _
            " AND " & Criteria4 & ")"
            End If

            If Criteria5 <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= " SERVICE_ID IN (SELECT EM.SERVICE_ID FROM EQUIPMENT_MOVEMENTS EM,EQUIPMENTS EQ,SITES S" & _
            " WHERE EM.EQUIPMENT_ID=EQ.EQUIPMENT_ID(+) AND EM.SITE_ID=S.SITE_ID(+) AND EM.MOVEMENT_TYPE=3" & _
            " AND " & Criteria5 & ")"
            End If

            If Cri <> "" Then
                If Criteria <> "" Then SQL &= " AND " Else SQL &= " WHERE "
                SQL &= " SERVICE_ID IN  (SELECT SERVICE_ID FROM V_SERVICE_ACTIONS WHERE " & Cri & ")"
            End If

            If RowLimit <> "" Then
                If Criteria <> "" Then SQL &= " AND "
                SQL &= "ROWNUM <= " & RowLimit
            End If

            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchServiceSite(Optional ByVal ServiceDateF As String = "", Optional ByVal ServiceDateT As String = "", _
                                      Optional ByVal ProjectType As String = "", Optional ByVal SiteGroupID As String = "", _
                                      Optional ByVal SiteName As String = "", Optional ByVal SiteID As String = "", _
                                      Optional ByVal MoveMentType As String = "", Optional ByVal InstallID As Integer = 0) As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = "", Criteria2 As String = "", Criteria3 As String = "" _
        , Criteria4 As String = "", Criteria5 As String = ""

        Try

            DB.AddCriteria(Criteria, "MOVEMENT_TYPE", MoveMentType, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteria(Criteria, "UPPER(SERVICE_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "DATE_INSTALLED", AppDateValue(ServiceDateF), AppDateValue(ServiceDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)

            If InstallID > 0 Then DB.AddCriteria(Criteria, "Install_Id", InstallID, DBUTIL.FieldTypes.ftNumeric)

            SQL = " SELECT  SALE_AREA, SALE_AREA_NAME, SITE_ID, SITE_NAME, STATUS, SALES_DISTRICT, SITE_STATUS_DESC, SITE_TYPE, SITE_TYPE_DESC, PROJECT_TYPE, " & _
                    " PROJECT_TYPE_DESC, PERCENT_COMPLETE, FTOTAL_SERVICE_TIME2, WORKSHEET, NGV_SITE " & _
                    " FROM  V_SERVICE_BOM_INSTALL3 "
            If Criteria <> "" Then SQL &= " WHERE " & Criteria



            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function ManageService(ByVal op As Integer, Optional ByRef ServiceID As String = Nothing _
    , Optional ByVal ServiceDate As String = Nothing, Optional ByVal ProjectType As String = Nothing _
    , Optional ByVal SiteGroupID As String = Nothing, Optional ByVal SiteID As String = Nothing _
    , Optional ByVal InformTitleID As String = Nothing, Optional ByVal InformName As String = Nothing _
    , Optional ByVal InformTel As String = Nothing, Optional ByVal InformEmail As String = Nothing _
    , Optional ByVal CallDetail As String = Nothing, Optional ByVal ProbCate As String = Nothing _
    , Optional ByVal ProbItem As String = Nothing, Optional ByVal SLAProfile As String = Nothing _
    , Optional ByVal SeverityLevel As String = Nothing, Optional ByVal ServiceStatus As String = Nothing _
    , Optional ByVal ServiceType As String = Nothing, Optional ByVal IssueAction As String = Nothing _
    , Optional ByVal VendorSLA As String = Nothing, Optional ByVal IssueSLA As String = Nothing _
    , Optional ByVal UpdateCloseDate As Boolean = False, Optional ByVal CloseDate As String = Nothing _
    , Optional ByVal TotalServiceTime As String = Nothing, Optional ByVal VendorAction As String = Nothing _
    , Optional ByVal RequireDate As String = Nothing, Optional ByVal CallMethod As String = Nothing _
    , Optional ByVal CallMethodOther As String = Nothing, Optional ByVal RootCause As String = Nothing _
    , Optional ByVal CloseReason As String = Nothing, Optional ByVal PendingReason As String = Nothing _
    , Optional ByVal RejectReason As String = Nothing, Optional ByVal ResolvedReason As String = Nothing _
    , Optional ByVal RefCallNumber As String = Nothing) As String

        Dim SQL1, SQL2, SQL, ServiceNo As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If ServiceID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    ServiceID = GenerateID("SERVICES", "SERVICE_ID") & ""
                    ServiceNo = GenerateID("SERVICES", "SERVICE_NO", GeneratePrefix(ProjectType), 4)
                    DB.AddSQL(op, SQL1, SQL2, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "SERVICE_NO", ServiceNo, DBUTIL.FieldTypes.ftText)
                    DB.AddSQL(op, SQL1, SQL2, "SERVICE_STATUS", "1", DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL2(op, SQL1, SQL2, "USER_CREATED", HttpContext.Current.Session("USER_NAME"), DBUTIL.FieldTypes.ftText)
                    DB.AddSQL2(op, SQL1, SQL2, "DATE_CREATED", Now, DBUTIL.FieldTypes.ftDateTime)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_DATE", AppDateValue(ServiceDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_GROUP_ID", SiteGroupID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_TITLE_ID", InformTitleID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_NAME", InformName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_TEL", InformTel, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_EMAIL", InformEmail, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "CALL_DETAIL", CallDetail, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "REQUIRE_DATE", AppDateValue(RequireDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "CALL_METHOD", CallMethod, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "CALL_METHOD_OTHER", CallMethodOther, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM_CATEGORY", ProbCate, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM_ITEM", ProbItem, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfile, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_TYPE", ServiceType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ISSUE_ACTION", IssueAction, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_ACTION", VendorAction, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_SLA", VendorSLA, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ISSUE_SLA", IssueSLA, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "TOTAL_SERVICE_TIME", TotalServiceTime, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ROOT_CAUSE", RootCause, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "CLOSE_REASON", CloseReason, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PENDING_REASON", PendingReason, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "REJECT_REASON", RejectReason, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "RESOLVED_REASON", ResolvedReason, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "REF_CALL_NUMBER", RefCallNumber, DBUTIL.FieldTypes.ftText) 'เพิ่ม 30/09/2010
                If UpdateCloseDate Then
                    DB.AddSQL2(op, SQL1, SQL2, "CLOSE_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddSQL2(op, SQL1, SQL2, "CLOSE_DATE", AppDateValue(CloseDate), DBUTIL.FieldTypes.ftDateTime)
                End If
            End If
            SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICES", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            If op = opINSERT Then
                ServiceID = ""
            End If
            Throw ex
        End Try

    End Function

    'Public Function ManageServiceAction(ByVal op As Integer, ByVal ServiceID As String _
    ', ByRef ActionID As String, Optional ByVal ActionDate As String = Nothing _
    ', Optional ByVal ResponseDate As String = Nothing, Optional ByVal UserName As String = Nothing _
    ', Optional ByVal ServiceStatus As String = Nothing, Optional ByVal AssignToGrp As String = Nothing _
    ', Optional ByVal AssignTo As String = Nothing, Optional ByVal SLAProfileID As String = Nothing _
    ', Optional ByVal SeverityLevel As String = Nothing, Optional ByVal Note As String = Nothing _
    ', Optional ByVal UpdateAction As Boolean = False, Optional ByVal UpdateResponse As Boolean = False _
    ', Optional ByVal UpdateResolvedDate As Boolean = False, Optional ByVal RefActionID As String = Nothing _
    ', Optional ByVal ResolvedDate As String = Nothing _
    ', Optional ByVal AssignToActionID As String = Nothing, Optional ByVal SLA As String = Nothing _
    ', Optional ByVal UpdateStatusDate As Boolean = False, Optional ByVal ActionTime As String = Nothing) As String

    '    Dim SQL1, SQL2, SQL As String
    '    Dim Criteria As String = ""

    '    Try
    '        SQL = "" : SQL1 = "" : SQL2 = ""

    '        If op = DBUTIL.opDELETE Then
    '            DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddCriteria(Criteria, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
    '        Else
    '            If ActionID <> "" Then
    '                op = DBUTIL.opUPDATE
    '                DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddCriteria(Criteria, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
    '            Else
    '                op = DBUTIL.opINSERT
    '                ActionID = GenerateID("SERVICE_ACTIONS", "ACTION_ID", usrCriteria:="SERVICE_ID=" + ServiceID) & ""
    '                DB.AddSQL(op, SQL1, SQL2, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddSQL(op, SQL1, SQL2, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
    '            End If

    '            If UpdateAction Then
    '                DB.AddSQL(op, SQL1, SQL2, "ACTION_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
    '            Else
    '                DB.AddSQL2(op, SQL1, SQL2, "ACTION_DATE", AppDateValue(ActionDate), DBUTIL.FieldTypes.ftDateTime)
    '            End If

    '            If UpdateResponse Then
    '                DB.AddSQL(op, SQL1, SQL2, "RESPONSE_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
    '            Else
    '                DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_DATE", AppDateValue(ResponseDate), DBUTIL.FieldTypes.ftDateTime)
    '            End If

    '            If UpdateResolvedDate Then
    '                DB.AddSQL(op, SQL1, SQL2, "RESOLVED_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
    '            Else
    '                DB.AddSQL2(op, SQL1, SQL2, "RESOLVED_DATE", AppDateValue(ResolvedDate), DBUTIL.FieldTypes.ftDateTime)
    '            End If

    '            If UpdateStatusDate Then
    '                DB.AddSQL(op, SQL1, SQL2, "STATUS_UPDATE_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
    '            End If

    '            DB.AddSQL2(op, SQL1, SQL2, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
    '            DB.AddSQL2(op, SQL1, SQL2, "SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO_GRP", AssignToGrp, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO", AssignTo, DBUTIL.FieldTypes.ftText)
    '            DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO_ACTION_ID", AssignToActionID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '            'DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_SLA_PROFILE_ID", AssignSLAProfileID, DBUTIL.FieldTypes.ftNumeric)
    '            'DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_SEVERITY_LEVEL", AssignSeverityLevel, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "REF_ACTION_ID", RefActionID, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "NOTE", Note, DBUTIL.FieldTypes.ftText)
    '            'DB.AddSQL2(op, SQL1, SQL2, "OLD_SERVICE_STATUS", OldServiceStatus, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "SLA", SLA, DBUTIL.FieldTypes.ftNumeric)
    '            DB.AddSQL2(op, SQL1, SQL2, "ACTION_TIME", ActionTime, DBUTIL.FieldTypes.ftNumeric)
    '        End If

    '        SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICE_ACTIONS", Criteria, True)
    '        DB.ExecSQL(SQL)
    '        Return ""
    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    'Public Function ManageServiceAction(ByVal op As Integer, ByVal ServiceID As String _
    ', ByRef ActionID As String, Optional ByVal ActionDate As String = Nothing _
    ', Optional ByVal ResponseDate As String = Nothing, Optional ByVal UserName As String = Nothing _
    ', Optional ByVal ServiceStatus As String = Nothing, Optional ByVal AssignToGrp As String = Nothing _
    ', Optional ByVal AssignTo As String = Nothing, Optional ByVal SLAProfileID As String = Nothing _
    ', Optional ByVal SeverityLevel As String = Nothing, Optional ByVal Note As String = Nothing _
    ', Optional ByVal UpdateAction As Boolean = False, Optional ByVal UpdateResponse As Boolean = False _
    ', Optional ByVal UpdateResolvedDate As Boolean = False, Optional ByVal RefActionID As String = Nothing _
    ', Optional ByVal ResolvedDate As String = Nothing, Optional ByVal OldServiceStatus As String = Nothing _
    ', Optional ByVal AssignSLAProfileID As String = Nothing, Optional ByVal AssignSeverityLevel As String = Nothing _
    ', Optional ByVal AssignToActionID As String = Nothing, Optional ByVal SLA As String = Nothing _
    ', Optional ByVal UpdateStatusDate As Boolean = False, Optional ByVal ActionTime As String = Nothing) As String


    Public Function ManageServiceBomInstall(ByVal op As Integer, ByRef InstalledId As Integer _
    , ByVal SiteId As String, ByVal BomGroupId As Integer, ByVal BomDetailId As Integer, ByVal dateInstalled As String, _
    ByVal dateCompleted As String, ByVal qty As Integer, ByVal systemStatus As Integer, ByVal moveMentType As Integer, _
    ByVal UpdateDate As String, ByVal UserUpdate As String) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""
        Dim InstallDate As String = SQLDate(AppDateValue(dateInstalled))
        Dim Completedate As String = SQLDate(AppDateValue(dateCompleted))

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""



            If op = DBUTIL.opINSERT Then
                InstalledId = GenerateID("SERVICE_BOM_INSTALL", "INSTALL_ID") & ""

                SQL1 = "INSTALL_ID ,SITE_ID,BOM_GROUP_ID ,BOM_DETAIL_ID,DATE_INSTALLED,DATE_COMPLETED ,"
                SQL1 &= "QTY, SYSTEM_STATUS, MOVEMENT_TYPE"

                SQL2 = InstalledId & " ,'" & SiteId & "', " & BomGroupId & " ," & BomDetailId & "," & InstallDate & "," & Completedate & ", "
                SQL2 &= qty & " , " & systemStatus & ", " & moveMentType '& "," & AppDateValue(UpdateDate) & ",'" & UserUpdate & "'"


            ElseIf op = DBUTIL.opUPDATE Then
                DB.AddCriteria(Criteria, "INSTALL_ID", InstalledId, DBUTIL.FieldTypes.ftNumeric)



                SQL1 = " SITE_ID= '" & SiteId & "',"
                SQL1 &= "BOM_GROUP_ID= " & BomGroupId & ", "
                SQL1 &= "BOM_DETAIL_ID= " & BomDetailId & ", "
                SQL1 &= "DATE_INSTALLED= " & InstallDate & ", "
                SQL1 &= "DATE_COMPLETED= " & Completedate & ", "
                SQL1 &= "QTY= " & qty & ", "
                SQL1 &= "SYSTEM_STATUS= " & systemStatus & ", "
                SQL1 &= "MOVEMENT_TYPE= " & moveMentType & ""
                'SQL1 &= "DATE_UPDATED= " & AppDateValue(UpdateDate) & ","
                'SQL1 &= "USER_UPDATED= '" & UserUpdate & "'"


            End If






            SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICE_BOM_INSTALL", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Public Function ManageServiceAction(ByVal op As Integer, ByVal ServiceID As String _
    , ByRef ActionID As String, Optional ByVal ActionDate As String = Nothing _
    , Optional ByVal ResponseDate As String = Nothing, Optional ByVal UserName As String = Nothing _
    , Optional ByVal ServiceStatus As String = Nothing, Optional ByVal AssignToGrp As String = Nothing _
    , Optional ByVal AssignTo As String = Nothing, Optional ByVal SLAProfileID As String = Nothing _
    , Optional ByVal SeverityLevel As String = Nothing, Optional ByVal Note As String = Nothing _
    , Optional ByVal UpdateAction As Boolean = False, Optional ByVal UpdateResponse As Boolean = False _
    , Optional ByVal UpdateResolvedDate As Boolean = False, Optional ByVal RefActionID As String = Nothing _
    , Optional ByVal ResolvedDate As String = Nothing _
    , Optional ByVal AssignSLAProfileID As String = Nothing, Optional ByVal AssignSeverityLevel As String = Nothing _
    , Optional ByVal AssignToActionID As String = Nothing, Optional ByVal SLA As String = Nothing _
    , Optional ByVal IsUpdateStatusDate As Boolean = False, Optional ByVal ActionTime As String = Nothing _
    , Optional ByVal ServiceStatusUpdate As String = Nothing, Optional ByVal AssignBy As String = Nothing _
    , Optional ByVal UpdateStatusDate As String = Nothing, Optional ByVal AssignToActionID2 As String = Nothing _
    , Optional ByVal Reason As String = Nothing, Optional ByVal RootCause As String = Nothing _
    , Optional ByVal UserGroupID As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If ActionID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    ActionID = GenerateID("SERVICE_ACTIONS", "ACTION_ID", usrCriteria:="SERVICE_ID=" + ServiceID) & ""
                    DB.AddSQL(op, SQL1, SQL2, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "ACTION_ID", ActionID, DBUTIL.FieldTypes.ftNumeric)
                End If

                If UpdateAction Then
                    DB.AddSQL(op, SQL1, SQL2, "ACTION_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddSQL2(op, SQL1, SQL2, "ACTION_DATE", AppDateValue(ActionDate), DBUTIL.FieldTypes.ftDateTime)
                End If

                If UpdateResponse Then
                    DB.AddSQL(op, SQL1, SQL2, "RESPONSE_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddSQL2(op, SQL1, SQL2, "RESPONSE_DATE", AppDateValue(ResponseDate), DBUTIL.FieldTypes.ftDateTime)
                End If

                If UpdateResolvedDate Then
                    DB.AddSQL(op, SQL1, SQL2, "RESOLVED_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddSQL2(op, SQL1, SQL2, "RESOLVED_DATE", AppDateValue(ResolvedDate), DBUTIL.FieldTypes.ftDateTime)
                End If

                If IsUpdateStatusDate Then
                    DB.AddSQL(op, SQL1, SQL2, "STATUS_UPDATE_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                Else
                    DB.AddSQL2(op, SQL1, SQL2, "STATUS_UPDATE_DATE", AppDateValue(UpdateStatusDate), DBUTIL.FieldTypes.ftDateTime)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "USER_GROUP_ID", UserGroupID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_STATUS", ServiceStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO_GRP", AssignToGrp, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO", AssignTo, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO_ACTION_ID", AssignToActionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_TO_ACTION_ID2", AssignToActionID2, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SLA_PROFILE_ID", SLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SEVERITY_LEVEL", SeverityLevel, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_SLA_PROFILE_ID", AssignSLAProfileID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_SEVERITY_LEVEL", AssignSeverityLevel, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "REF_ACTION_ID", RefActionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "NOTE", Note, DBUTIL.FieldTypes.ftText)
                'DB.AddSQL2(op, SQL1, SQL2, "OLD_SERVICE_STATUS", OldServiceStatus, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SLA", SLA, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ACTION_TIME", ActionTime, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SERVICE_STATUS_UPDATE", ServiceStatusUpdate, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "ASSIGN_BY", AssignBy, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "REASON", Reason, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ROOT_CAUSE", RootCause, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICE_ACTIONS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function ManageServiceAttachment(ByVal op As Integer, ByVal ServiceID As String _
    , ByRef AttachmentID As String, Optional ByVal AttachmentDesc As String = Nothing _
    , Optional ByVal AttachmentFile As String = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "ATTACHMENT_ID", AttachmentID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If AttachmentID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "ATTACHMENT_ID", AttachmentID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    AttachmentID = GenerateID("SERVICE_ATTACHMENTS", "ATTACHMENT_ID", usrCriteria:="SERVICE_ID=" + ServiceID) & ""
                    DB.AddSQL(op, SQL1, SQL2, "SERVICE_ID", ServiceID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "ATTACHMENT_ID", AttachmentID, DBUTIL.FieldTypes.ftNumeric)
                End If
                DB.AddSQL2(op, SQL1, SQL2, "ATTACHMENT_DESC", AttachmentDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ATTACHMENT_FILE", AttachmentFile, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "SERVICE_ATTACHMENTS", Criteria, True)
            DB.ExecSQL(SQL)
            Return ""
        Catch ex As Exception
            Throw ex
            AttachmentID = ""
        End Try

    End Function
#End Region

#Region "Import"
    Public Function SearchImportLog(Optional ByVal ImportID As String = "" _
    , Optional ByVal ProjectType As String = "", Optional ByVal Importor As String = "" _
    , Optional ByVal VendorName As String = "", Optional ByVal ImportDay As String = "" _
    , Optional ByVal ImportMonth As String = "", Optional ByVal ImportYear As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "IM.IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.IMPORT_DAY", ImportDay, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.IMPORT_MONTH", ImportMonth, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.IMPORT_YEAR", ImportYear, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddCriteriaRange(Criteria, "IM.IMPORT_DATE", AppDateValue(ImportDateF), AppDateValue(ImportDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(IM.USER_UPDATED)", Importor.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(VD.VENDOR_NAME)", VendorName.ToUpper, DBUTIL.FieldTypes.ftText)

            SQL = "SELECT IM.*,VD.VENDOR_NAME,PT.PROJECT_TYPE_DESC" & _
            ",DECODE(IM.IMPORT_STATUS,'P','Processing','C','Confirmed',NULL) AS IMPORT_STATUS_DESC " & _
            " FROM IMPORT_LOGS IM,VENDORS VD,REF_PROJECT_TYPES PT " & _
            " WHERE IM.VENDOR_CODE=VD.VENDOR_CODE(+) AND IM.PROJECT_TYPE=PT.PROJECT_TYPE(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY IM.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchReconcile(Optional ByVal VendorName As String = "", Optional ByVal ProjectType As String = "" _
    , Optional ByVal VSF As String = "", Optional ByVal VST As String = "" _
    , Optional ByVal VSSystemF As String = "", Optional ByVal VSSystemT As String = "" _
    , Optional ByVal VSNoShowSysF As String = "", Optional ByVal VSNoShowSysT As String = "" _
    , Optional ByVal SNoShowVendF As String = "", Optional ByVal SNoShowVendT As String = "" _
    , Optional ByVal Day As String = "", Optional ByVal Month As String = "" _
    , Optional ByVal Year As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "VR.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(VD.VENDOR_NAME)", VendorName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "VR.VENDOR_SERVICE_CNT", VSF, VST, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "VR.SERVICE_CNT", VSSystemF, VSSystemT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "VR.VS_NO_SHOW_SYSTEM_CNT", VSNoShowSysF, VSNoShowSysT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteriaRange(Criteria, "VR.V_SERVICE_NO_SHOW_CNT", SNoShowVendF, SNoShowVendT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "VR.IMPORT_DAY", Day, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "VR.IMPORT_MONTH", Month, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "VR.IMPORT_YEAR", Year, DBUTIL.FieldTypes.ftNumeric)


            SQL = "SELECT VR.*,VD.VENDOR_NAME FROM V_RECONCILES VR,VENDORS VD WHERE VR.VENDOR_CODE=VD.VENDOR_CODE"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY VD.VENDOR_NAME"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchImportDetailTemp(Optional ByVal ProjectType As String = "", Optional ByVal VendorName As String = "" _
    , Optional ByVal ServiceNo As String = "", Optional ByVal OpenDateF As String = "" _
    , Optional ByVal OpenDateT As String = "", Optional ByVal InDateF As String = "", Optional ByVal InDateT As String = "" _
    , Optional ByVal FinishDateF As String = "", Optional ByVal FinishDateT As String = "", Optional ByVal Status As String = "" _
    , Optional ByVal SeverityF As String = "", Optional ByVal SeverityT As String = "", Optional ByVal SiteID As String = "" _
    , Optional ByVal SiteName As String = "", Optional ByVal InformerBy As String = "", Optional ByVal Problem As String = "" _
    , Optional ByVal ProblemType As String = "", Optional ByVal VendorCode As String = "" _
    , Optional ByVal ImportID As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "IM.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(IM.VENDOR_CODE)", VendorCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(VD.VENDOR_NAME)", VendorName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.JOB_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            'DB.AddCriteriaRange(Criteria, "IM.IMPORT_DATE", ImportDateF, ImportDateT, DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.OPEN_DATE", AppDateValue(OpenDateF), AppDateValue(OpenDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.IN_DATE", AppDateValue(InDateF), AppDateValue(InDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.FINISH_DATE", AppDateValue(FinishDateF), AppDateValue(FinishDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(IMD.STATUS)", Status.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "IMD.SEVERITY_LEVEL", SeverityF, SeverityT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(IMD.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.INFORMER_NAME)", InformerBy.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.PROBLEM)", Problem.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.PROBLEM_TYPE)", ProblemType.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "IM.IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT IMD.*,IM.VENDOR_CODE,IM.PROJECT_TYPE,IM.IMPORT_DAY,IM.IMPORT_MONTH,IM.IMPORT_YEAR" & _
            ",VD.VENDOR_NAME FROM IMPORT_DETAIL_TEMPS IMD,IMPORT_LOGS IM,VENDORS VD" & _
            " WHERE IMD.IMPORT_ID=IM.IMPORT_ID AND IM.VENDOR_CODE=VD.VENDOR_CODE(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY IM.DATE_UPDATED DESC"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchImportDetail(Optional ByVal ProjectType As String = "", Optional ByVal VendorName As String = "" _
    , Optional ByVal ServiceNo As String = "", Optional ByVal OpenDateF As String = "" _
    , Optional ByVal OpenDateT As String = "", Optional ByVal InDateF As String = "", Optional ByVal InDateT As String = "" _
    , Optional ByVal FinishDateF As String = "", Optional ByVal FinishDateT As String = "", Optional ByVal Status As String = "" _
    , Optional ByVal SeverityF As String = "", Optional ByVal SeverityT As String = "", Optional ByVal SiteID As String = "" _
    , Optional ByVal SiteName As String = "", Optional ByVal InformerBy As String = "", Optional ByVal Problem As String = "" _
    , Optional ByVal ProblemType As String = "", Optional ByVal VendorCode As String = "" _
    , Optional ByVal ImportID As String = "", Optional ByVal ImportDay As String = "" _
    , Optional ByVal ImportMonth As String = "", Optional ByVal ImportYear As String = "" _
    , Optional ByVal OtherCriteria As String = "", Optional ByVal OrderBy As String = "") As DataTable
        Dim DT As DataTable = Nothing
        Dim SQL As String = "", Criteria As String = ""

        Try
            Criteria = OtherCriteria
            DB.AddCriteria(Criteria, "IM.PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(IM.VENDOR_CODE)", VendorCode.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "IM.IMPORT_DAY", ImportDay, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.IMPORT_MONTH", ImportMonth, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "IM.IMPORT_YEAR", ImportYear, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(VD.VENDOR_NAME)", VendorName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.JOB_NO)", ServiceNo.ToUpper, DBUTIL.FieldTypes.ftText)
            'DB.AddCriteriaRange(Criteria, "IM.IMPORT_DATE", ImportDateF, ImportDateT, DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.OPEN_DATE", AppDateValue(OpenDateF), AppDateValue(OpenDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.IN_DATE", AppDateValue(InDateF), AppDateValue(InDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteriaRange(Criteria, "IMD.FINISH_DATE", AppDateValue(FinishDateF), AppDateValue(FinishDateT), DBUTIL.FieldTypes.ftDate)
            DB.AddCriteria(Criteria, "UPPER(IMD.STATUS)", Status.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(Criteria, "IMD.SEVERITY_LEVEL", SeverityF, SeverityT, DBUTIL.FieldTypes.ftNumeric)
            DB.AddCriteria(Criteria, "UPPER(IMD.SITE_ID)", SiteID.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.SITE_NAME)", SiteName.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.INFORMER_NAME)", InformerBy.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.PROBLEM)", Problem.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "UPPER(IMD.PROBLEM_TYPE)", ProblemType.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteria(Criteria, "IM.IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT IMD.*,IM.IMPORT_DAY,IM.IMPORT_MONTH,IM.IMPORT_YEAR,IM.VENDOR_CODE" & _
            ",IM.PROJECT_TYPE,PT.PROJECT_TYPE_DESC,VD.VENDOR_NAME FROM IMPORT_DETAILS IMD,IMPORT_LOGS IM" & _
            ",VENDORS VD,REF_PROJECT_TYPES PT WHERE IMD.IMPORT_ID=IM.IMPORT_ID(+) " & _
            "AND IM.VENDOR_CODE=VD.VENDOR_CODE(+) AND IM.PROJECT_TYPE=PT.PROJECT_TYPE(+)"

            If Criteria <> "" Then SQL &= " AND " & Criteria
            If OrderBy <> "" Then
                SQL &= " ORDER BY " & OrderBy
            Else
                SQL &= " ORDER BY IM.IMPORT_YEAR DESC,IM.IMPORT_MONTH DESC,IM.IMPORT_DAY DESC,VD.VENDOR_NAME,IMD.JOB_NO"
            End If
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageImportLog(ByVal op As Integer, ByRef ImportID As String _
    , Optional ByVal ImportDay As String = Nothing, Optional ByVal ImportMonth As String = Nothing, Optional ByVal ImportYear As String = Nothing _
    , Optional ByVal ProjectType As String = Nothing, Optional ByVal VendorCode As String = Nothing _
    , Optional ByVal FileName As String = Nothing, Optional ByVal ImportFileName As String = Nothing, Optional ByVal TotalLine As String = Nothing _
    , Optional ByVal MatchLine As String = Nothing, Optional ByVal UnmatchLine As String = Nothing _
    , Optional ByVal UnmatchDetail As String = Nothing, Optional ByVal ImportStatus As String = Nothing _
    , Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
            Else
                If ImportID <> "" Then
                    op = DBUTIL.opUPDATE
                    DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = DBUTIL.opINSERT
                    ImportID = GenerateID("IMPORT_LOGS", "IMPORT_ID") & ""
                    DB.AddSQL(op, SQL1, SQL2, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                    'DB.AddSQL(op, SQL1, SQL2, "IMPORT_DATE", Now, DBUTIL.FieldTypes.ftDateTime)
                End If
                DB.AddSQL2(op, SQL1, SQL2, "IMPORT_DAY", ImportDay, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "IMPORT_MONTH", ImportMonth, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "IMPORT_YEAR", ImportYear, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "PROJECT_TYPE", ProjectType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "VENDOR_CODE", VendorCode, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "FILE_NAME", FileName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "IMPORT_FILE_NAME", ImportFileName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "TOTAL_LINES", TotalLine, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "MATCH_LINES", MatchLine, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "UNMATCH_LINES", UnmatchLine, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "UNMATCH_DETAIL", UnmatchDetail, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "IMPORT_STATUS", ImportStatus, DBUTIL.FieldTypes.ftText)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "IMPORT_LOGS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            Throw ex
            ImportID = ""
        End Try

    End Function

    Public Function ManageImportDetail(ByVal op As Integer, ByVal ImportID As String _
    , ByVal JobNo As String, Optional ByVal OpenDate As String = Nothing, Optional ByVal InDate As String = Nothing _
    , Optional ByVal FinishDate As String = Nothing, Optional ByVal Status As String = Nothing, Optional ByVal ETC As String = Nothing _
    , Optional ByVal Severity As String = Nothing, Optional ByVal SiteID As String = Nothing, Optional ByVal SiteName As String = Nothing _
    , Optional ByVal InformerBy As String = Nothing, Optional ByVal UserName As String = Nothing, Optional ByVal Point As String = Nothing _
    , Optional ByVal Problem As String = Nothing, Optional ByVal Other As String = Nothing, Optional ByVal DownTime As String = Nothing _
    , Optional ByVal Solved As String = Nothing, Optional ByVal System As String = Nothing, Optional ByVal ExpirySYS As String = Nothing _
    , Optional ByVal WO As String = Nothing, Optional ByVal ProblemType As String = Nothing, Optional ByVal Line As String = Nothing _
    , Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
            Else
                If op = opINSERT Then
                    DB.AddSQL(op, SQL1, SQL2, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
                Else
                    DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "OPEN_DATE", AppDateValue(OpenDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "IN_DATE", AppDateValue(InDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "FINISH_DATE", AppDateValue(FinishDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "STATUS", Status, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ETC", AppDateValue(ETC), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "SEVERITY_LEVEL", Severity, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_NAME", SiteName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_NAME", InformerBy, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "POINT", Point, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM", Problem, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "OTHER", Other, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DOWN_TIME", DownTime, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SOLVED", Solved, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_NAME", System, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EXPIRY_SYS", ExpirySYS, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "WO_NO", WO, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM_TYPE", ProblemType, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LINE", Line, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "IMPORT_DETAILS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Public Function ManageImportDetailTemp(ByVal op As Integer, ByVal ImportID As String _
    , ByVal JobNo As String, Optional ByVal OpenDate As String = Nothing, Optional ByVal InDate As String = Nothing _
    , Optional ByVal FinishDate As String = Nothing, Optional ByVal Status As String = Nothing, Optional ByVal ETC As String = Nothing _
    , Optional ByVal Severity As String = Nothing, Optional ByVal SiteID As String = Nothing, Optional ByVal SiteName As String = Nothing _
    , Optional ByVal InformerBy As String = Nothing, Optional ByVal UserName As String = Nothing, Optional ByVal Point As String = Nothing _
    , Optional ByVal Problem As String = Nothing, Optional ByVal Other As String = Nothing, Optional ByVal DownTime As String = Nothing _
    , Optional ByVal Solved As String = Nothing, Optional ByVal System As String = Nothing, Optional ByVal ExpirySYS As String = Nothing _
    , Optional ByVal WO As String = Nothing, Optional ByVal ProblemType As String = Nothing, Optional ByVal Line As String = Nothing _
    , Optional ByVal Conn As OleDb.OleDbConnection = Nothing, Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As String

        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""

            If op = DBUTIL.opDELETE Then
                DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
            Else
                If op = opINSERT Then
                    DB.AddSQL(op, SQL1, SQL2, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
                Else
                    DB.AddCriteria(Criteria, "IMPORT_ID", ImportID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddCriteria(Criteria, "JOB_NO", JobNo, DBUTIL.FieldTypes.ftText)
                End If

                DB.AddSQL2(op, SQL1, SQL2, "OPEN_DATE", AppDateValue(OpenDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "IN_DATE", AppDateValue(InDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "FINISH_DATE", AppDateValue(FinishDate), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "STATUS", Status, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "ETC", AppDateValue(ETC), DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL2(op, SQL1, SQL2, "SEVERITY_LEVEL", Severity, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_ID", SiteID, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SITE_NAME", SiteName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "INFORMER_NAME", InformerBy, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "USER_NAME", UserName, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "POINT", Point, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM", Problem, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "OTHER", Other, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "DOWN_TIME", DownTime, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SOLVED", Solved, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "SYSTEM_NAME", System, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "EXPIRY_SYS", ExpirySYS, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "WO_NO", WO, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "PROBLEM_TYPE", ProblemType, DBUTIL.FieldTypes.ftText)
                DB.AddSQL2(op, SQL1, SQL2, "LINE", Line, DBUTIL.FieldTypes.ftNumeric)
            End If

            SQL = DB.CombineSQL(op, SQL1, SQL2, "IMPORT_DETAIL_TEMPS", Criteria, True)
            DB.ExecSQL(SQL, Conn, Trans)
            Return ""
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region
End Class
