Imports System.Globalization
Imports System.Data

Public Module DropdownController

    Sub New()

    End Sub

    Public Function GenDropDownCtrlJson(ByVal TbName As String, ByVal CtrlID As String, ByVal FieldValue As String, ByVal FieldDesc As String, ByVal key_id_val As String, Optional ByVal ValSelected As String = "") As String
        Dim ddlCtrl As String = "", CtrlName As String = ""
        Dim DT As DataTable = GetData(TbName, key_id_val)

        ddlCtrl = "{'Table' :" & vbCrLf
        ddlCtrl &= "[" & vbCrLf
        If TbName <> "V_EQUIPMENT_PROBLEM_RELATION" And TbName <> "REF_SEVERITY_LEVELS" Then
            Dim label As String = IIf(TbName = "REF_PROJECT_TYPES", "== เลือกทั้งหมด ==", "== เลือก ==")
            ddlCtrl &= "{"
            ddlCtrl &= String.Format("'value' : '','label' : '{0}'", label)
            ddlCtrl &= "}," & vbCrLf
        End If

        For Each DR In DT.Rows
            ddlCtrl &= "{"
            ddlCtrl &= String.Format("'value' : '{0}','label' : '{1}'", DR(FieldValue).ToString(), DR(FieldDesc).ToString())
            ddlCtrl &= "}," & vbCrLf
        Next
        ddlCtrl &= "]" & vbCrLf
        ddlCtrl &= "}" & vbCrLf
        Return ddlCtrl
    End Function

    Public Function GetData(ByVal TbName As String, ByVal key_id_val As String) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = ""
        Select Case TbName
            Case "REF_PROJECT_TYPES"
                sql = String.Format("SELECT PROJECT_TYPE, PROJECT_TYPE_DESC FROM REF_PROJECT_TYPES WHERE TO_CHAR(PROJECT_TYPE) LIKE '%{0}%' ORDER BY PROJECT_TYPE", key_id_val)
            Case "SITE_GROUP_LISTS"
                sql = String.Format( _
"SELECT SG.SITE_GROUP_ID, SG.SITE_GROUP_NAME " & _
"FROM SITE_GROUP_LISTS SGL " & _
"INNER JOIN SITE_GROUPS SG ON SGL.SITE_GROUP_ID = SG.SITE_GROUP_ID " & _
"INNER JOIN SITES SS ON SGL.SITE_ID = SS.SITE_ID  " & _
"WHERE 1=1  " & _
"AND TO_CHAR(SS.PROJECT_TYPE) LIKE '%{0}%' " & _
"GROUP BY SG.SITE_GROUP_ID, SG.SITE_GROUP_NAME  " & _
"ORDER BY SG.SITE_GROUP_ID " _
        , key_id_val _
        )
            Case "REF_SALE_AREAS"
                sql = String.Format( _
 "SELECT RSA.SALE_AREA, RSA.SALE_AREA_NAME  " & _
"FROM REF_SALE_AREAS RSA  " & _
"INNER JOIN SITES SS ON RSA.SALE_AREA = SS.SALE_AREA  " & _
"WHERE 1=1  " & _
"AND TO_CHAR(SS.PROJECT_TYPE) LIKE '%{0}%' " & _
"GROUP BY RSA.SALE_AREA, RSA.SALE_AREA_NAME  " & _
"ORDER BY RSA.SALE_AREA  " _
        , key_id_val _
        )
            Case "SLA_PROFILES"
                sql = String.Format( _
"SELECT SLA_PROFILE_ID, PROFILE_NAME AS PROFILE_NAME  " & _
"FROM SLA_PROFILES " & _
"WHERE 1=1 " & _
"AND ACTIVE_FLAG='Y'  " & _
"AND SLA_TYPE = 1  " & _
"AND TO_CHAR(PROJECT_TYPE) LIKE '%{0}%' " & _
"GROUP BY SLA_PROFILE_ID, PROFILE_NAME  " & _
"ORDER BY PROFILE_NAME  " _
        , key_id_val _
        )
            Case "V_EQUIPMENT_PROBLEM_RELATION"
                sql = String.Format( _
"SELECT EQUIPMENT_TYPE, EQUIPMENT_TYPE_DESC AS EQUIPMENT_TYPE_DESC " & _
"FROM V_EQUIPMENT_PROBLEM_RELATION  " & _
"WHERE 1=1  " & _
"AND TO_CHAR(PROJECT_TYPE) LIKE '%{0}%' " & _
"GROUP BY EQUIPMENT_TYPE, EQUIPMENT_TYPE_DESC  " & _
"ORDER BY EQUIPMENT_TYPE_DESC " _
        , key_id_val _
        )
            Case "REF_SEVERITY_LEVELS"
                sql = String.Format( _
"SELECT RSL.SEVERITY_LEVEL, RSL.SEVERITY_LEVEL_DESC " & _
"FROM REF_SEVERITY_LEVELS RSL " & _
"INNER JOIN SLA_DETAILS SD ON RSL.SEVERITY_LEVEL = SD.SEVERITY_LEVEL " & _
"INNER JOIN SLA_PROFILES SP ON SD.SLA_PROFILE_ID = SP.SLA_PROFILE_ID " & _
"WHERE 1=1 " & _
"AND SP.ACTIVE_FLAG = 'Y' " & _
"AND SP.SLA_TYPE = 1 " & _
"AND TO_CHAR(SP.PROJECT_TYPE) LIKE '%{0}%' " & _
"GROUP BY RSL.SEVERITY_LEVEL, RSL.SEVERITY_LEVEL_DESC " & _
"ORDER BY SEVERITY_LEVEL_DESC " _
        , key_id_val _
        )
        End Select

        If sql <> "" Then
            dt = DAL.QueryData(sql, Nothing, Nothing)
        End If

        Return dt
    End Function

End Module
