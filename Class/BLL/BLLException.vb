#Region ".NET Framework Class Import"
Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Collections
#End Region



<Serializable()> _
Public Class BLLException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal errorMessage As String)
        MyBase.New(errorMessage)
    End Sub

    Public Sub New(ByVal ex As Exception)
        MyBase.New(ex.Message, ex)
    End Sub

    Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New(info, context)
    End Sub

    Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.GetObjectData(info, context)
    End Sub
End Class

