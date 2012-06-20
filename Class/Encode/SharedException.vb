

#Region ".NET Framework Class Import"
Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Collections
Imports System.Web.Services.Protocols
#End Region

<Serializable()> _
Public Class SharedException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal errorMessage As String)
        MyBase.New(errorMessage)
    End Sub

    Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New(info, context)
    End Sub

    Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.GetObjectData(info, context)
    End Sub
End Class

Public Class ApplicationError : Inherits SoapHeader
    Public Message As String
End Class
