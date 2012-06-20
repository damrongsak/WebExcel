Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary


    Public Module SerializeHelper

        Public Function SerializeObject(ByVal pObject As Object) As String
            Dim MS As New MemoryStream
            Dim XS As XmlSerializer = New XmlSerializer(pObject.GetType())
            Dim XW As XmlTextWriter = New XmlTextWriter(MS, Encoding.UTF8)
            Dim UTF8 As New System.Text.UTF8Encoding
            Dim XmlString As String = Nothing

            XS.Serialize(XW, pObject)
            MS = CType(XW.BaseStream, MemoryStream)
            XmlString = UTF8.GetString(MS.ToArray())

            Return XmlString
        End Function

        Public Sub DeserializeObject(ByRef pObject As Object, ByVal XmlString As String)
            Dim UTF8 As New System.Text.UTF8Encoding
            Dim MS As New MemoryStream(UTF8.GetBytes(XmlString))
            Dim XS As XmlSerializer = New XmlSerializer(pObject.GetType())
            Dim XW As XmlTextWriter = New XmlTextWriter(MS, Encoding.UTF8)

            XS.Serialize(XW, pObject)
            MS = CType(XW.BaseStream, MemoryStream)
            XmlString = UTF8.GetString(MS.ToArray())
            pObject = XS.Deserialize(MS)
        End Sub

        Public Sub SerializeFile(ByVal pObject As Object, ByVal Filename As String)
            Dim XS As XmlSerializer = New XmlSerializer(pObject.GetType())
            Dim SW As StreamWriter = Nothing
            Dim MapFilename As String

            Try
                MapFilename = HttpContext.Current.Server.MapPath(Filename)
                SW = New StreamWriter(MapFilename)
                XS.Serialize(SW, pObject)
            Finally
                If Not IsNothing(SW) Then SW.Close()
                SW = Nothing
            End Try
        End Sub

        Public Sub DeserializeFile(ByRef pObject As Object, ByVal Filename As String)
            Dim XS As XmlSerializer = New XmlSerializer(pObject.GetType())
            Dim SR As StreamReader = Nothing
            Dim MapFilename As String

            Try
                MapFilename = HttpContext.Current.Server.MapPath(Filename)
                If IO.File.Exists(MapFilename) Then
                    SR = New StreamReader(MapFilename)
                    pObject = XS.Deserialize(SR)
                End If
            Finally
                If Not IsNothing(SR) Then SR.Close()
                SR = Nothing
            End Try
        End Sub

    End Module

