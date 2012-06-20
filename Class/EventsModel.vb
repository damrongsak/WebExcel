Namespace DataModel
    Public Class EventsModel

        Private _id As Integer
        Private _name As String
        Private _eventstart As Date
        Private _eventend As Date
        Sub New()

        End Sub

        Public Property ID() As Integer
            Get
                Return Me._id
            End Get
            Set(ByVal value As Integer)
                Me._id = value
            End Set
        End Property

        Public Property NAME() As String
            Get
                Return Me._name
            End Get
            Set(ByVal value As String)
                Me._name = value
            End Set
        End Property

        Public Property EVENTSTART() As Date
            Get
                Return Me._eventstart
            End Get
            Set(ByVal value As Date)
                Me._eventstart = value
            End Set
        End Property

        Public Property EVENTEND() As Date
            Get
                Return Me._eventend
            End Get
            Set(ByVal value As Date)
                Me._eventend = value
            End Set
        End Property


    End Class
End Namespace
