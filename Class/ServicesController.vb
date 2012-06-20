Imports System.Globalization
Imports System.Data

Public Class ServicesController

    Private _Time_Total As Integer = 0
    Private _dt_Service_List As DataTable

    Sub New()

    End Sub

    Public Function getServiceTime(ByVal sid As String) As Integer
        Dim cnt As Integer = 0
        Dim dr1 As DataRow
        Dim dr2 As DataRow
        cnt = _dt_Service_List.Rows.Count
        If cnt = 0 Then
            Me._dt_Service_List = Me.getServiceList(sid)
            cnt = Me._dt_Service_List.Rows.Count
        End If
        For i As Integer = 0 To cnt - 1
            dr1 = Me._dt_Service_List.Rows(i)
            If (i + 1) < cnt Then
                dr2 = Me._dt_Service_List.Rows(i + 1)
            Else
                dr2 = dr1
            End If
            If dr1("SERVICE_STATUS_DESC").ToString() <> "Closed" Then
                If dr1("SERVICE_STATUS_DESC").ToString() <> "Pending" Then
                    Dim d1 As DateTime = Date.ParseExact(dr1("ACTION_DATE").ToString(), "yyyy-MM-dd HH:mm:ss", New CultureInfo("en-US"))
                    Dim d2 As DateTime = Date.ParseExact(dr2("ACTION_DATE").ToString(), "yyyy-MM-dd HH:mm:ss", New CultureInfo("en-US"))
                    Me._Time_Total = Me._Time_Total + Me.getDiffTime(d1, d2)
                End If
            End If
        Next
        Me._dt_Service_List.Dispose()

        Return Me._Time_Total
    End Function

    Public Function getServiceList(ByVal sid As String) As DataTable
        Dim dt As DataTable = New DataTable()
        dt = DAL.QueryData(String.Format("SELECT * FROM V_SERVICES_TEMP WHERE SERVICE_ID ='{0}'", sid), Nothing, Nothing)
        Return dt
    End Function

    Public Function getDiffTime(ByVal d1 As DateTime, ByVal d2 As DateTime) As Integer
        Dim total As Integer = 0
        total = DateDiff(DateInterval.Minute, d1, d2)
        Return total
    End Function
End Class
