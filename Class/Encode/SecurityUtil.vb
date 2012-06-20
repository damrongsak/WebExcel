#Region ".NET Framework Class Import"
Imports System.Web.Security
Imports System.Security
Imports System.Security.Principal
Imports System
Imports System.DirectoryServices
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
#End Region

Public Class SecurityUtil

#Region "Internal member variables"
    Private _appName As String
    Private _appKey As String
    Private _keyString As String
    Private _key() As Byte
    Private _LDAP_Path As String = ""
    Private _LDAP_FilterAttribute As String = ""
#End Region

    Private Const symmProvider As String = "TripleDESCryptoServiceProvider"
    Private Const hashProvider As String = "MD5CryptoServiceProvider"

    Private Function SymDecrypt(ByVal EncryptedString As String, ByVal KeyString As String) As String
        Dim key() As Byte
        Dim DecryptedString As String

        Try
            CryptoUtility.SymmetricCryptographer.Initialize()
            key = Encoding.Unicode.GetBytes(KeyString)
            DecryptedString = CryptoUtility.SymmetricCryptographer.DecryptString(EncryptedString, key)
            CryptoUtility.SymmetricCryptographer.Unload(Nothing, Nothing)
            Return DecryptedString
        Catch ex As Exception
            CryptoUtility.SymmetricCryptographer.Unload(Nothing, Nothing)
            Throw New SharedException("Decryption error. " + ex.Message)
        End Try
    End Function

    Public Function DecryptData(ByVal EncryptedText As String) As String
        Dim Key As String

        Try
            Key = ConfigurationManager.AppSettings("SecurityKey") & ""
            Return SymDecrypt(EncryptedText, Key)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function IsADAuthenticated(ByVal Domain As String, ByVal UserName As String, ByVal Password As String) As Boolean
        Dim entry As DirectoryEntry
        Dim search As DirectorySearcher
        Dim result As SearchResult
        Dim obj As Object
        Dim ADPersonalID As String = ""


        Try
            If Domain = "" Then ' Demo mode
                If UserName = Password Then
                    Return True
                Else
                    Return False
                End If
            Else
                ADPersonalID = System.Configuration.ConfigurationManager.AppSettings("ADPersonalID") & ""
                If _LDAP_Path = "" Then _LDAP_Path = "LDAP://" & Domain
                entry = New DirectoryEntry(_LDAP_Path, UserName, Password)
                obj = entry.NativeObject
                search = New DirectorySearcher(entry)
                search.Filter = "(&(objectClass=user)(sAMAccountName=" + UserName + "))" '"{SAMAccountName=" + UserName + ")"
                search.PropertiesToLoad.Add("cn")
                search.PropertiesToLoad.Add(ADPersonalID)
                result = search.FindOne
                entry = Nothing

                If Not IsNothing(result) Then
                    _LDAP_Path = result.Path
                    _LDAP_FilterAttribute = CType(result.Properties("cn")(0), String)
                    'Find ID Card
                    For Each result In search.FindAll()
                        HttpContext.Current.Session("ID_CARD") = GetProperty(result, ADPersonalID) & ""
                    Next
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            entry = Nothing
            IsADAuthenticated = False
            'Throw New SharedException("Error authenticating user. " + ex.Message)
        End Try
    End Function

    Public Sub QueryAD(ByVal Domain As String, ByVal UserName As String, ByVal Password As String, Optional ByRef UserDesc As String = "", Optional ByRef PositionName As String = "", Optional ByRef Department As String = "", Optional ByRef Mail As String = "", Optional ByRef TelNo As String = "")
        Dim entry As DirectoryEntry
        Dim search As DirectorySearcher
        Dim result As SearchResult
        Dim obj As Object

        Try
            If Domain = "" Then ' Demo mode
                'If UserName = Password Then
                '    Return True
                'Else
                '    Return False
                'End If
            Else
                If _LDAP_Path = "" Then _LDAP_Path = "LDAP://" & Domain
                entry = New DirectoryEntry(_LDAP_Path, UserName, Password)
                obj = entry.NativeObject
                search = New DirectorySearcher(entry)
                search.Filter = "(&(objectClass=user)(sAMAccountName=" + UserName + "))" '"{SAMAccountName=" + UserName + ")"
                search.PropertiesToLoad.Add("cn")
                search.PropertiesToLoad.Add("displayName")
                search.PropertiesToLoad.Add("department")
                search.PropertiesToLoad.Add("title")
                search.PropertiesToLoad.Add("mail")
                search.PropertiesToLoad.Add("telephoneNumber")
                search.PropertiesToLoad.Add("homephone")
                search.PropertiesToLoad.Add("mobile")
                search.CacheResults = True
                entry = Nothing

                For Each result In search.FindAll()
                    UserDesc = GetProperty(result, "displayName") & ""
                    PositionName = GetProperty(result, "title") & ""
                    Department = GetProperty(result, "department") & ""
                    Mail = GetProperty(result, "mail") & ""
                    TelNo = GetProperty(result, "telephoneNumber") & ""
                    If TelNo <> "" Then
                        TelNo += "," & GetProperty(result, "homephone") & ""
                    Else
                        TelNo = GetProperty(result, "homephone") & ""
                    End If
                    If TelNo <> "" Then
                        TelNo += "," & GetProperty(result, "mobile") & ""
                    Else
                        TelNo = GetProperty(result, "mobile") & ""
                    End If
                Next
            End If

        Catch ex As Exception
            Dim msg As String

            msg = ex.Message
            entry = Nothing
            'QueryAD = False
            'Throw New SharedException("Error authenticating user. " + ex.Message)
        End Try
    End Sub

    Public Function GetProperty(ByVal result As SearchResult, ByVal PropertyName As String) As String
        Dim Value As String = ""

        Try
            If Not IsNothing(result) Then
                If result.Properties(PropertyName).Count > 0 Then
                    Value = result.Properties(PropertyName)(0) & ""
                End If
            End If
        Catch
        End Try

        Return Value
    End Function

End Class
