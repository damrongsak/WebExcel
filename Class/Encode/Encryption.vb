#Region ".NET Framework Class Import"
Imports System.Security.Cryptography
#End Region

Public Class Encryption

    Private Const Key As String = "@PTTEP!"

    Public Function EncryptData(ByVal value As String) As String
        Dim des As New TripleDESCryptoServiceProvider
        des.IV = New Byte(7) {}

        Dim pdb As New PasswordDeriveBytes(Key, New Byte(-1) {})
        des.Key = pdb.CryptDeriveKey("RC2", "MD5", 128, New Byte(7) {})

        Dim ms As New IO.MemoryStream((value.Length * 2) - 1)
        Dim encStream As New CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write)
        Dim plainBytes As Byte() = Text.Encoding.UTF8.GetBytes(value)

        encStream.Write(plainBytes, 0, plainBytes.Length)
        encStream.FlushFinalBlock()

        Dim encryptedBytes(CInt(ms.Length - 1)) As Byte

        ms.Position = 0
        ms.Read(encryptedBytes, 0, CInt(ms.Length))
        encStream.Close()

        Return Convert.ToBase64String(encryptedBytes)
    End Function

    Public Function DecryptData(ByVal value As String) As String
        Dim des As New TripleDESCryptoServiceProvider
        des.IV = New Byte(7) {}

        Dim pdb As New PasswordDeriveBytes(Key, New Byte(-1) {})
        des.Key = pdb.CryptDeriveKey("RC2", "MD5", 128, New Byte(7) {})

        Dim encryptedBytes As Byte() = Convert.FromBase64String(value)
        Dim ms As New IO.MemoryStream(value.Length)
        Dim decStream As New CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write)

        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        Dim plainBytes(CInt(ms.Length - 1)) As Byte

        ms.Position = 0
        ms.Read(plainBytes, 0, CInt(ms.Length))
        decStream.Close()

        Return Text.Encoding.UTF8.GetString(plainBytes)
    End Function


End Class

