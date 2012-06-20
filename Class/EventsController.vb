Imports System.Globalization
Imports System.Data
Imports PTTICT.ServiceTracking.USL.Web.DataModel

Public Class EventsController

    Sub New()

    End Sub

    Public Sub ActionCreate(ByVal data As EventsModel)
        Dim sql As String = String.Format("INSERT INTO EVENT(NAME, EVENTSTART, EVENTEND) VALUES('{0}',TO_DATE('{1}','DD-MM-YYYY HH24:MI:SS'),TO_DATE('{2}','DD-MM-YYYY HH24:MI:SS'))" _
                                          , data.NAME _
                                          , data.EVENTSTART.ToString("dd-MM-yyyy HH:mm:ss", New CultureInfo("en-US")) _
                                          , data.EVENTEND.ToString("dd-MM-yyyy HH:mm:ss", New CultureInfo("en-US")) _
                                          )
        DAL.ExecSQL(sql, Nothing, Nothing)
    End Sub

    Public Sub ActionUpdate(ByVal id As Integer, ByVal data As EventsModel)
        data.NAME = IIf(String.IsNullOrEmpty(data.NAME), "", String.Format("name = '{0}',", data.NAME))
        Dim sql As String = String.Format("UPDATE EVENT SET {0} eventstart=TO_DATE('{1}','DD-MM-YYYY HH24:MI:SS'), eventend=TO_DATE('{2}','DD-MM-YYYY HH24:MI:SS') WHERE id = {3}" _
                                          , data.NAME _
                                          , data.EVENTSTART.ToString("dd-MM-yyyy HH:mm:ss", New CultureInfo("en-US")) _
                                          , data.EVENTEND.ToString("dd-MM-yyyy HH:mm:ss", New CultureInfo("en-US")) _
                                          , id _
                                          )
        DAL.ExecSQL(sql, Nothing, Nothing)
    End Sub

    Public Sub ActionDelete(ByVal id As Integer)
        DAL.ExecSQL(String.Format("DELETE EVENT WHERE ID={0}", id), Nothing, Nothing)
    End Sub

    Public Function ActionView(ByVal id As Integer) As DataTable
        Dim dt As DataTable = New DataTable()
        dt = DAL.QueryData(String.Format("SELECT * FROM EVENT WHERE id = {0}", id), Nothing, Nothing)
        Return dt
    End Function

    Public Function ActionView() As DataTable
        Dim dt As DataTable = New DataTable()
        dt = DAL.QueryData(String.Format("SELECT * FROM EVENT"), Nothing, Nothing)
        Return dt
    End Function
End Class
