#Region ".NET Framework Class Import"
Imports System
Imports System.Xml.Serialization
Imports System.Runtime.Serialization
Imports System.Collections
#End Region



<Serializable()> _
Public Class AuditData

#Region "Internal member variables"
    Private _TransactionDate As Date
    Private _UserId As String
    Private _DeptName As String
    Private _Category As String
    Private _Action As String
    Private _MachineName As String
#End Region

    Public Sub New(ByVal TransactionDate As Date, _
                   ByVal UserId As String, _
                   ByVal DeptName As String, _
                   ByVal Category As String, _
                   ByVal Action As String, _
                   ByVal MachineName As String)
        _TransactionDate = TransactionDate
        _UserId = UserId
        _DeptName = DeptName
        _Category = Category
        _Action = Action
        _MachineName = MachineName
    End Sub

    Public Sub New()
        _TransactionDate = Nothing
        _UserId = Nothing
        _DeptName = Nothing
        _Category = Nothing
        _Action = Nothing
        _MachineName = Nothing
    End Sub

#Region "Property"

    Public Property TransactionDate() As Date
        Get
            Return _TransactionDate
        End Get
        Set(ByVal Value As Date)
            _TransactionDate = Value
        End Set
    End Property

    Public Property UserId() As String
        Get
            Return _UserId
        End Get
        Set(ByVal Value As String)
            _UserId = Value
        End Set
    End Property

    Public Property DeptName() As String
        Get
            Return _DeptName
        End Get
        Set(ByVal Value As String)
            _DeptName = Value
        End Set
    End Property

    Public Property Category() As String
        Get
            Return _Category
        End Get
        Set(ByVal Value As String)
            _Category = Value
        End Set
    End Property

    Public Property Action() As String
        Get
            Return _Action
        End Get
        Set(ByVal Value As String)
            _Action = Value
        End Set
    End Property

    Public Property MachineName() As String
        Get
            Return _MachineName
        End Get
        Set(ByVal Value As String)
            _MachineName = Value
        End Set
    End Property

#End Region

End Class


<Serializable()> _
Public Class AuditDatas
    Inherits CollectionBase

    Default Public ReadOnly Property Item(ByVal index As Integer) As AuditData
        Get
            If (index < 0 Or index >= Me.InnerList.Count) Then
                Throw New Exception("Index has to be between 0 and " & (Me.InnerList.Count - 1).ToString())
            Else
                Return CType(Me.InnerList(index), AuditData)
            End If
        End Get
    End Property

    Public Sub Add(ByVal info As AuditData)
        Me.InnerList.Add(info)
    End Sub

    Public Sub SetItem(ByVal index As Integer, ByVal value As AuditData)
        If (index < 0 Or index >= Me.InnerList.Count) Then
            Throw New Exception("Index has to be between 0 and " & (Me.InnerList.Count - 1).ToString())
        Else
            Me.InnerList(index) = value
        End If
    End Sub

End Class

