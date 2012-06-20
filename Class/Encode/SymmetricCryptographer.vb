
'============================================================================================================
' CryptoUtility.SymmetricCryptographer
'	
' SymmetricCryptographer.cs
'
' PURPOSE:		To encrypt and decrypt sensitive data.
' DATE:		03.03.2004
' AUTHORS:		mstuart
' 
'	<SAMPLE CODE>
' 
'============================================================================================================
' 
'============================================================================================================
Imports System
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Collections
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Text


Namespace CryptoUtility
    ''' <summary>
    ''' This class provides symmetric encryption services using the Rijndael Managed base class library provider.
    ''' It is modelled after the eWeek/Microsoft "OpenHack4" competition code.
    ''' 
    ''' 
    ''' </summary>
    <ComVisible(False)> _
               Friend MustInherit Class SymmetricCryptographer

#Region "Constructors"


        ''' <summary>
        ''' cctor sets up the symmetric algorithm instance and the RNG provider
        ''' </summary>
        Shared Sub New()
            Initialize()
        End Sub 'New


#End Region

#Region "Declarations"

        Private Shared _rng As RNGCryptoServiceProvider = Nothing
        Private Shared _alg As SymmetricAlgorithm = Nothing

        Private Shared _hasBeenInitialized As Boolean = False

        Private Const IV_SIZE As Integer = 16
        Private Const KEY_SIZE_BYTES As Integer = 32

#End Region


#If DEBUG Then
        ''' <summary>
        ''' Allows check to see if it has been initialized.
        ''' Only used during testing.
        ''' </summary>

        Friend Shared ReadOnly Property IsInitialized() As Boolean
            Get
                Return _hasBeenInitialized
            End Get
        End Property
#End If
        ''' <summary>
        ''' internal method used to set up symmetric algorithm with given key
        ''' </summary>
        Friend Shared Sub Initialize()
            '  check if already initialized if so return
            If _hasBeenInitialized Then
                Return
            End If
            '  assume init fails until end
            _hasBeenInitialized = False

            '  new up the RNG
            _rng = New RNGCryptoServiceProvider

            ' Use SymmetricCrypto static factory method on Rijndael derived class to create new Rijndael instance
            _alg = RijndaelManaged.Create()

            '  set key size to passed key value
            _alg.KeySize = KEY_SIZE_BYTES * 8

            '  trap unload events so we can clean up algorithm via Dispose()
            AddHandler AppDomain.CurrentDomain.DomainUnload, AddressOf SymmetricCryptographer.Unload
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf SymmetricCryptographer.Unload

            '  set init to true since we got here
            _hasBeenInitialized = True
        End Sub 'Initialize



        ''' <summary>
        ''' This method returns base64-encoded CipherText from ClearText
        ''' </summary>
        ''' <param name="clearString">ClearText string</param>
        ''' <param name="key">byte array full cryptographic key</param>
        ''' <returns>base64-encoded CipherText</returns>
        Friend Shared Function EncryptString(ByVal clearString As String, ByVal key() As Byte) As String
            '  check if we've been initalized, if we have not, throw:
            If Not _hasBeenInitialized Then
                Throw New ApplicationException("Symmetric cryptographer has not been initialized.")
            End If
            Dim clearText As Byte() = Encoding.Unicode.GetBytes(clearString)
            'Dim key As Byte() = Encoding.Unicode.GetBytes(keyString)
            Dim data As Byte() = Nothing
            Dim output As Byte() = Nothing
            Dim newIV As Byte() = Nothing

            ' generate 16 random bytes with RNG, use that as IV
            newIV = New Byte(IV_SIZE - 1) {}
            SyncLock _rng
                _rng.GetBytes(newIV)
            End SyncLock
            '  get encryptor, set the IV in ctor; "using" is because it's disposable
            Dim trans As ICryptoTransform = _alg.CreateEncryptor(key, newIV)
            Try
                Dim ms As New MemoryStream
                Try
                    Dim cs As New CryptoStream(ms, trans, CryptoStreamMode.Write)
                    Try
                        cs.Write(clearText, 0, clearText.Length)
                        cs.FlushFinalBlock()
                        data = ms.ToArray()
                    Finally
                        cs.Close()
                    End Try
                Finally
                    ms.Close()
                End Try
            Finally
                trans.Dispose()
            End Try

            ' now append the IV to the beginning of the ciphered text
            output = New Byte(IV_SIZE + data.Length - 1) {}
            Buffer.BlockCopy(newIV, 0, output, 0, newIV.Length)
            Buffer.BlockCopy(data, 0, output, IV_SIZE, data.Length)

            Return Convert.ToBase64String(output)
        End Function 'EncryptString


        ''' <summary>
        ''' This method returns ClearText from base64-encoded CipherText
        ''' </summary>
        ''' <param name="cipherString">Base64-encoded CipherText string</param>
        ''' <param name="key">byte array full cryptographic key</param>
        ''' <returns>ClearText string</returns>
        Friend Shared Function DecryptString(ByVal cipherString As String, ByVal key() As Byte) As String
            '  check if we've been initalized, if we have not, throw:
            If Not _hasBeenInitialized Then
                Throw New ApplicationException("Symmetric cryptographer has not been initialized.")
            End If
            Dim cipherBlob As Byte() = Convert.FromBase64String(cipherString)
            Dim cipherText(cipherBlob.Length - IV_SIZE - 1) As Byte
            Dim data As Byte() = Nothing
            Dim iv(IV_SIZE - 1) As Byte

            'BUG MSTUART 06.17.2004:  since we appended the IV, of course this will never fail.  
            '  move this check to AFTER stripping off IV.
            '  check block size against length of input string; can't be greater
            Dim blockSize As Integer = _alg.BlockSize / 8

            If cipherBlob.Length < blockSize Then
                Throw New ArgumentException("Data length must be greater than block size.")
            End If
            '  strip salt (IV) off first 16 bytes of input
            Buffer.BlockCopy(cipherBlob, 0, iv, 0, IV_SIZE)
            '  put actual ciphertext back into cipherText array
            Buffer.BlockCopy(cipherBlob, IV_SIZE, cipherText, 0, cipherBlob.Length - IV_SIZE)

            Dim trans As ICryptoTransform = _alg.CreateDecryptor(key, iv)
            Try
                Dim ms As New MemoryStream
                Try
                    Dim cs As New CryptoStream(ms, trans, CryptoStreamMode.Write)
                    Try
                        cs.Write(cipherText, 0, cipherText.Length)
                        cs.FlushFinalBlock()
                        data = ms.ToArray()
                    Finally
                        cs.Close()
                    End Try
                Finally
                    ms.Close()
                End Try
            Finally
                trans.Dispose()
            End Try

            Return Encoding.Unicode.GetString(data)
        End Function 'DecryptString

        ''' <summary>
        ''' Internal overload, sinks the ProcessExit and AppDomain Unload events so that we 
        ''' get first chance to clean up cryptObj objects--which are all IDisposable.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Friend Shared Sub Unload(ByVal sender As Object, ByVal e As EventArgs)
            If Not _hasBeenInitialized Then
                Return
            End If
            CType(_alg, IDisposable).Dispose()

            _hasBeenInitialized = False
        End Sub 'Unload
    End Class 'SymmetricCryptographer 
End Namespace 'CryptoUtility 

