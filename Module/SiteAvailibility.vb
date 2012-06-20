Imports System.Globalization

Public Class SiteAvailibility
    Private DAL As New DALComponent
    Private _AvbRows As New DataTable()
    Private _Date_Start As Date = Date.Now
    Private _Date_End As Date = Date.Now
    Private _Sale_Area As String = ""
    Private _Project_type As String = ""
    Private _Cul As CultureInfo = New CultureInfo("en-US")
    Private _Site_ID As String
    Private _SiteGroup As String
    Private _MovementType As String

    Sub New()
        Me._AvbRows.Columns.Add("SALE_AREA", GetType(String))
        Me._AvbRows.Columns.Add("SALE_AREA_NAME", GetType(String))
        Me._AvbRows.Columns.Add("SERVICE_TIME", GetType(String))
        Me._AvbRows.Columns.Add("CNT_SERVICES", GetType(Integer))
        Me._AvbRows.Columns.Add("CNT_SITES", GetType(Integer))
        Me._AvbRows.Columns.Add("PERCENT", GetType(Double))
    End Sub

    Private Property AvbRows() As DataTable
        Get
            Return Me._AvbRows
        End Get
        Set(ByVal value As DataTable)
            Me._AvbRows = value
        End Set
    End Property

    Public Property DateStart() As Date
        Get
            Return Me._Date_Start
        End Get
        Set(ByVal value As Date)
            Me._Date_Start = value
        End Set
    End Property

    Public Property DateEnd() As Date
        Get
            Return Me._Date_End
        End Get
        Set(ByVal value As Date)
            Me._Date_End = value
        End Set
    End Property

    Public Property SaleArea() As String
        Get
            Return Me._Sale_Area
        End Get
        Set(ByVal value As String)
            Me._Sale_Area = value
        End Set
    End Property

    Public Property ProjectType() As String
        Get
            Return Me._Project_type
        End Get
        Set(ByVal value As String)
            Me._Project_type = value
        End Set
    End Property

    Public Property Site_ID() As String
        Get
            Return Me._Site_ID
        End Get
        Set(ByVal value As String)
            Me._Site_ID = value
        End Set
    End Property

    Public Property SiteGroup() As String
        Get
            Return Me._SiteGroup
        End Get
        Set(ByVal value As String)
            Me._SiteGroup = value
        End Set
    End Property

    Public Property MovementType() As String
        Get
            Return Me._MovementType
        End Get
        Set(ByVal value As String)
            Me._MovementType = value
        End Set
    End Property

    Private Function getArea() As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = String.Format( _
            "SELECT SITES.SALE_AREA, REF_SALE_AREAS.SALE_AREA_NAME, count(SITES.SITE_ID) AS CNT_SITES " & _
            "FROM SITES INNER JOIN " & _
            "REF_SITE_TYPES ON SITES.SITE_TYPE = REF_SITE_TYPES.SITE_TYPE INNER JOIN " & _
            "REF_PROJECT_TYPES ON SITES.PROJECT_TYPE = REF_PROJECT_TYPES.PROJECT_TYPE RIGHT OUTER JOIN " & _
            "REF_SALE_AREAS ON SITES.SALE_AREA = REF_SALE_AREAS.SALE_AREA " & _
            "WHERE (SITES.STATUS = 1) AND (SITES.PROJECT_TYPE = 2) " & _
            "GROUP BY SITES.SALE_AREA, REF_SALE_AREAS.SALE_AREA_NAME " _
            )
        dt = DAL.QueryData(sql, Nothing, Nothing)
        Return dt
    End Function

    Private Function getServicesBySaleArea() As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = ""

        sql = String.Format( _
"SELECT SALE_AREA, SALE_AREA_NAME, SUM(FALL_SERVICE_TIME) AS SERVICE_TIME, COUNT(SERVICE_ID) AS CNT_SERVICES " & _
"FROM V_SERVICES_AVAILABILITY " & _
"WHERE (SEVERITY_LEVEL = 1) AND  " & _
"(SERVICE_TYPE <> 3) AND  " & _
"(SERVICE_DATE BETWEEN TO_DATE('{0}', 'DD/MM/YYYY') AND TO_DATE('{1} 23:59:59', 'DD/MM/YYYY HH24:MI:SS')) " & _
"AND (PROJECT_TYPE = {2}) AND (SALE_AREA LIKE '%{3}%') AND (TO_CHAR(SITE_GROUP_ID) LIKE '%{4}%')" & _
"GROUP BY SALE_AREA, SALE_AREA_NAME " _
            , Me._Date_Start.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Date_End.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Project_type _
            , Me._Sale_Area _
            , Me.SiteGroup _
            )
        dt = DAL.QueryData(sql, Nothing, Nothing)
        Return dt
    End Function

    Private Function getServicesByEquipment() As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = String.Format( _
            "SELECT SITE_ID, SITE_NAME, PROBLEM_TYPE,  PROBLEM_TYPE_DESC, TO_CHAR(SUM(FALL_SERVICE_TIME)) AS SERVICE_TIME, COUNT(SERVICE_ID) AS CNT_SERVICES, 0 AS PERCENT, 1 AS CNT_SITES " & _
            "FROM V_SERVICES_AVAILABILITY " & _
            "WHERE(SEVERITY_LEVEL = 1) AND (SERVICE_DATE BETWEEN TO_DATE('{0}', 'DD/MM/YYYY') AND TO_DATE('{1} 23:59:59', 'DD/MM/YYYY HH24:MI:SS'))  " & _
            "AND (SERVICE_TYPE <> 3) " & _
            "GROUP BY SERVICE_TYPE, SITE_ID, SITE_NAME, PROJECT_TYPE, PROBLEM_TYPE,  PROBLEM_TYPE_DESC " & _
            "HAVING (SITE_ID = '{2}') AND (PROJECT_TYPE = {3})" _
            , Me._Date_Start.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Date_End.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Site_ID _
            , Me._Project_type _
            )
        dt = DAL.QueryData(sql, Nothing, Nothing)
        Return dt
    End Function

    Private Function getServicesBySite() As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = String.Format( _
            "SELECT SALE_AREA, SALE_AREA_NAME, SITE_ID, SITE_NAME, TO_CHAR(SUM(FALL_SERVICE_TIME)) AS SERVICE_TIME, COUNT(SERVICE_ID) AS CNT_SERVICES, 0 AS PERCENT, 1 AS CNT_SITES " & _
            "FROM V_SERVICES_AVAILABILITY " & _
            "WHERE(SEVERITY_LEVEL = 1) AND (SERVICE_DATE BETWEEN TO_DATE('{0}', 'DD/MM/YYYY') AND TO_DATE('{1} 23:59:59', 'DD/MM/YYYY HH24:MI:SS'))  " & _
            "AND (SERVICE_TYPE <> 3) " & _
            "GROUP BY SALE_AREA, SALE_AREA_NAME, SERVICE_TYPE, SITE_ID, SITE_NAME, PROJECT_TYPE " & _
            "HAVING (SALE_AREA = '{2}') AND (PROJECT_TYPE = {3})" _
            , Me._Date_Start.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Date_End.ToString("dd/MM/yyyy", Me._Cul) _
            , Me._Sale_Area _
            , Me._Project_type _
            )
        dt = DAL.QueryData(sql, Nothing, Nothing)
        Return dt
    End Function


    Public Function getEquipmentBySite() As DataTable
        Dim dt As DataTable = New DataTable()
        Dim sql As String = String.Format( _
            "SELECT  V_EQUIPMENTS_INSTALL.SITE_ID, V_EQUIPMENTS_INSTALL.SITE_NAME, V_EQUIPMENTS_INSTALL.EQUIPMENT_ID, V_EQUIPMENTS_INSTALL.SERIAL_NO, " & _
            "V_EQUIPMENTS_INSTALL.EQUIPMENT_DESC, V_EQUIPMENTS_INSTALL.QUANTITY, V_EQUIPMENTS_INSTALL.EQUIPMENT_TYPE,  " & _
            "V_EQUIPMENTS_INSTALL.EQUIPMENT_TYPE_DESC, V_EQUIPMENTS_INSTALL.EQUIPMENT_STATUS, V_EQUIPMENTS_INSTALL.EQUIPMENT_STATUS_DESC,  " & _
            "REF_PROBLEM_TYPES.PROBLEM_TYPE, REF_PROBLEM_TYPES.PROBLEM_TYPE_DESC " & _
            "FROM   REF_PROBLEM_TYPES RIGHT OUTER JOIN " & _
            "REF_EQUIPMENT_TYPES ON REF_PROBLEM_TYPES.EQUIPMENT_TYPE = REF_EQUIPMENT_TYPES.EQUIPMENT_TYPE RIGHT OUTER JOIN " & _
            "V_EQUIPMENTS_INSTALL ON REF_EQUIPMENT_TYPES.EQUIPMENT_TYPE = V_EQUIPMENTS_INSTALL.EQUIPMENT_TYPE " & _
            "WHERE (V_EQUIPMENTS_INSTALL.SITE_ID = '{0}') AND V_EQUIPMENTS_INSTALL.EQUIPMENT_STATUS = 1" & _
            "ORDER BY V_EQUIPMENTS_INSTALL.EQUIPMENT_TYPE " _
            , Me._Site_ID _
            )
        dt = DAL.QueryData(sql, Nothing, Nothing)
        Return dt
    End Function

    Public Function CalAreaAvailibility() As DataTable
        Dim dtSite As DataTable = New DataTable()
        Dim dtServices As DataTable = New DataTable()
        dtSite = Me.getArea()
        dtServices = Me.getServicesBySaleArea()

        Dim tbl As DataTable
        Dim colsA() As String = {"SALE_AREA", "SALE_AREA_NAME", "CNT_SITES"}
        Dim colsB() As String = {"SERVICE_TIME", "CNT_SERVICES"}
        Dim sKey As String = "SALE_AREA"

        tbl = MergeData(dtSite, dtServices, colsA, colsB, sKey)

        For Each dr As DataRow In tbl.Rows
            Dim sale_area As String = dr("SALE_AREA").ToString()
            Dim sale_area_name As String = dr("SALE_AREA_NAME").ToString()
            Dim xx As String = dr("CNT_SITES").ToString().Trim()
            Dim yy As String = dr("SERVICE_TIME").ToString().Trim()
            Dim zz As String = dr("CNT_SERVICES").ToString().Trim()
            xx = IIf(IsNumeric(xx), xx, 0)
            yy = IIf(IsNumeric(yy), yy, 0)
            zz = IIf(IsNumeric(zz), zz, 0)
            Dim cnt_sites As Double = Convert.ToDouble(xx)
            Dim service_time As Double = Convert.ToDouble(yy)
            Dim cnt_services As Double = Convert.ToDouble(zz)
            Dim percent As Double = 0
            Dim y As Double = Me._Date_Start.Year()
            Dim m As Double = Me._Date_Start.Month()
            Dim days As Double = System.DateTime.DaysInMonth(y, m)

            If (Me._Date_Start <> Me._Date_End) Then
                days = Convert.ToDouble(DateDiff(DateInterval.DayOfYear, Me._Date_Start, Me._Date_End))
            End If

            percent = 100 - (service_time / ((cnt_sites * (24 * 60) * days) * 100))
            percent = Convert.ToDouble(String.Format("{0:n3}", percent))

            Dim _dr As DataRow = Me._AvbRows.NewRow()
            _dr("SALE_AREA") = sale_area
            _dr("SALE_AREA_NAME") = sale_area_name
            _dr("CNT_SITES") = cnt_sites
            _dr("SERVICE_TIME") = MinuteToString(service_time)
            _dr("CNT_SERVICES") = cnt_services
            _dr("PERCENT") = percent
            Me._AvbRows.Rows.Add(_dr)
        Next

        Return Me._AvbRows
    End Function

    Public Function CalSiteAvailibility()
        Dim dt As DataTable = Me.getServicesBySite()
        Dim dt2 As DataTable = dt.Clone()
        Dim i As Integer = 0
        For i = 0 To dt.Rows.Count - 1
            Dim dr As DataRow = dt.Rows(i)
            Dim sale_area As String = dr("SALE_AREA").ToString()
            Dim sale_area_name As String = dr("SALE_AREA_NAME").ToString()
            Dim site_id As String = dr("SITE_ID").ToString()
            Dim site_name As String = dr("SITE_NAME").ToString()
            Dim xx As String = dr("CNT_SITES").ToString().Trim()
            Dim yy As String = dr("SERVICE_TIME").ToString().Trim()
            Dim zz As String = dr("CNT_SERVICES").ToString().Trim()
            xx = IIf(IsNumeric(xx), xx, 0)
            yy = IIf(IsNumeric(yy), yy, 0)
            zz = IIf(IsNumeric(zz), zz, 0)
            Dim cnt_sites As Double = Convert.ToDouble(xx)
            Dim service_time As Double = Convert.ToDouble(yy)
            Dim cnt_services As Double = Convert.ToDouble(zz)
            Dim percent As Double = 0
            Dim y As Double = Me._Date_Start.Year()
            Dim m As Double = Me._Date_Start.Month()
            Dim days As Double = System.DateTime.DaysInMonth(y, m)
            If (Me._Date_Start <> Me._Date_End) Then
                days = Convert.ToDouble(DateDiff(DateInterval.DayOfYear, Me._Date_Start, Me._Date_End))
            End If

            percent = 100 - (service_time / ((cnt_sites * (24 * 60) * days) * 100))
            percent = Convert.ToDouble(String.Format("{0:n3}", percent))
            Dim dr2 As DataRow = dt2.NewRow
            dr2("SALE_AREA") = sale_area
            dr2("SALE_AREA_NAME") = sale_area_name
            dr2("SITE_ID") = site_id
            dr2("SITE_NAME") = site_name
            dr2("CNT_SITES") = cnt_sites
            dr2("SERVICE_TIME") = MinuteToString(service_time)
            dr2("CNT_SERVICES") = cnt_services
            dr2("CNT_SITES") = 1
            dr2("PERCENT") = percent
            dt2.Rows.Add(dr2)
        Next
        Return dt2
    End Function

    Public Function CalEquipmentBySite()
        Dim dt As DataTable = Me.getEquipmentBySite()
        Dim dt2 As DataTable = Me.getServicesByEquipment()
        Dim tbl As DataTable
        Dim colsA() As String = {"SITE_ID", "SITE_NAME", "EQUIPMENT_ID", "SERIAL_NO", "EQUIPMENT_DESC", "QUANTITY", "EQUIPMENT_TYPE", "EQUIPMENT_TYPE_DESC", "EQUIPMENT_STATUS", "EQUIPMENT_STATUS_DESC", "PROBLEM_TYPE", "PROBLEM_TYPE_DESC"}
        Dim colsB() As String = {"SERVICE_TIME", "CNT_SERVICES", "PERCENT"}
        Dim sKey As String = "PROBLEM_TYPE"
        tbl = MergeData(dt, dt2, colsA, colsB, sKey)
        Return tbl
    End Function

    Function MinuteToString(ByVal minutes As Integer)
        Dim ts As New TimeSpan(0, minutes, 0)
        Dim s As String = String.Format("{0} ชั่วโมง {1} นาที ", ts.Hours, ts.Minutes, ts.Seconds)
        Return s
    End Function
    'Dim tbl As DataTable
    'Dim colsA() As String = {"ProjectNo", "Quantity"}
    'Dim colsB() As String = {"Customer"}
    'Dim sKey As String = "ProjectNo"
    'tbl = MergeData(tblA, tblB, colsA, colsB, sKey)
    Private Function MergeData(ByVal tblA As DataTable, ByVal tblB As DataTable, _
                              ByVal colsA() As String, ByVal colsB() As String, _
                              ByVal sKey As String) As DataTable

        Dim tbl As DataTable
        Dim col As DataColumn
        Dim sColumnName As String
        Dim row As DataRow
        Dim newRow As DataRow
        Dim dv As DataView

        tbl = New DataTable
        dv = tblB.DefaultView

        For Each sColumnName In colsA
            col = tblA.Columns(sColumnName)
            tbl.Columns.Add(New DataColumn(col.ColumnName, col.DataType))
        Next
        For Each sColumnName In colsB
            col = tblB.Columns(sColumnName)
            tbl.Columns.Add(New DataColumn(col.ColumnName, col.DataType))
        Next

        For Each row In tblA.Rows
            newRow = tbl.NewRow
            For Each sColumnName In colsA
                newRow(sColumnName) = row(sColumnName)
            Next
            Try
                dv.RowFilter = (sKey & " = " & row(sKey).ToString)
                If dv.Count = 1 Then
                    For Each sColumnName In colsB
                        newRow(sColumnName) = dv(0).Item(sColumnName)
                    Next
                End If

            Catch ex As Exception

            End Try
            tbl.Rows.Add(newRow)
        Next

        Return tbl

    End Function
End Class
