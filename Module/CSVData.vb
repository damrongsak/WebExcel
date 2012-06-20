'******************************************************************************
'*                                                                            *
'*                      Comma Separated Value Data Class                      *
'*                                                                            *
'*                             By John Priestley                              *
'*                                                                            *
'******************************************************************************
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text.RegularExpressions

Public Class CSVData
    Implements IDisposable

    Dim dsCSV As DataSet
    Dim mSeparator As Char = ","
    Dim mTextQualifier As Char = """"
    Dim mData() As String
    Dim mHeader As Boolean

    Private regQuote As New Regex("^(\x22)(.*)(\x22)(\s*,)(.*)$", RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
    Private regNormal As New Regex("^([^,]*)(\s*,)(.*)$", RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
    Private regQuoteLast As New Regex("^(\x22)([\x22*]{2,})(\x22)$", RegexOptions.IgnoreCase)
    Private regNormalLast As New Regex("^.*$", RegexOptions.IgnoreCase)

    Protected Disposed As Boolean

#Region " Load CSV "
    '
    ' Load CSV
    '
    Public Sub LoadCSV(ByVal CSVFile As String)
        LoadCSV(CSVFile, False)
    End Sub

    '
    ' Load CSV - Has Header
    '
    Public Sub LoadCSV(ByVal CSVFile As String, ByVal HasHeader As Boolean)
        Dim sr As New StreamReader(CSVFile, Encoding.GetEncoding(874))
        Dim idx As Integer
        Dim CntFirstLine As Integer = 0
        Dim bFirstLine As Boolean = True
        Dim dr As DataRow
        Try
            mHeader = HasHeader
            SetupRegEx()

            If File.Exists(CSVFile) = False Then
                Throw New Exception(CSVFile & " does not exist.")
            End If

            If Not dsCSV Is Nothing Then
                dsCSV.Clear()
                dsCSV.Tables.Clear()
                dsCSV.Dispose()
                dsCSV = Nothing
            End If

            dsCSV = New DataSet("CSV")
            dsCSV.Tables.Add("CSVData")

            Do While sr.Peek > -1
                ProcessLine(sr.ReadLine())

                '
                ' Create Columns
                '

                If Not IsNothing(mData) Then
                    If bFirstLine = True Then
                        CntFirstLine = mData.GetUpperBound(0) + 1
                        For idx = 0 To mData.GetUpperBound(0)
                            If mHeader = True Then
                                dsCSV.Tables("CSVData").Columns.Add(mData(idx), GetType(String))
                            Else
                                dsCSV.Tables("CSVData").Columns.Add("Column" & idx, GetType(String))
                            End If
                        Next
                    End If

                    '
                    ' Add Data
                    '
                    If Not (bFirstLine = True And mHeader = True) Then
                        dr = dsCSV.Tables("CSVData").NewRow()

                        If CntFirstLine < mData.GetUpperBound(0) + 1 Then
                            dsCSV.Tables("CSVData").Columns.Add("Column" & mData.GetUpperBound(0), GetType(String))
                            CntFirstLine = mData.GetUpperBound(0) + 1
                        End If

                        For idx = 0 To mData.GetUpperBound(0)
                            Try
                                dr(idx) = mData(idx)
                            Catch ex As Exception

                            End Try
                        Next

                        dsCSV.Tables("CSVData").Rows.Add(dr)
                        dsCSV.AcceptChanges()
                    End If

                    bFirstLine = False
                End If
            Loop

            sr.Close()
        Catch ex As Exception
            sr.Close()
            Throw ex
        End Try
    End Sub

    '
    ' Load CSV with custom separator
    '
    Public Sub LoadCSV(ByVal CSVFile As String, ByVal Separator As Char)
        LoadCSV(CSVFile, Separator, False)
    End Sub

    '
    ' Load CSV with custom separator and Has Header
    '
    Public Sub LoadCSV(ByVal CSVFile As String, ByVal Separator As Char, ByVal HasHeader As Boolean)
        mSeparator = Separator
        Try
            LoadCSV(CSVFile, HasHeader)
        Catch ex As Exception
            Throw New Exception("CSV Error", ex)
        End Try
    End Sub

    '
    ' Load CSV with custom separator and text qualifier
    '
    Public Sub LoadCSV(ByVal CSVFile As String, ByVal Separator As Char, ByVal TxtQualifier As Char)
        LoadCSV(CSVFile, Separator, TxtQualifier, False)
    End Sub

    '
    ' Load CSV with custom separator and text qualifier
    '
    Public Sub LoadCSV(ByVal CSVFile As String, ByVal Separator As Char, ByVal TxtQualifier As Char, ByVal HasHeader As Boolean)
        mSeparator = Separator
        mTextQualifier = TxtQualifier
        Try
            LoadCSV(CSVFile, HasHeader)
        Catch ex As Exception
            Throw New Exception("CSV Error", ex)
        End Try
    End Sub
#End Region
#Region " Process Line "
    '
    ' Process Line
    '
    Private Sub ProcessLine(ByVal sLine As String)
        Dim sData As String
        Dim iSep As Integer = 0
        Dim iQuote As String = ""
        Dim m As Match
        Dim idx As Integer
        Dim mc As MatchCollection

        Erase mData
        sLine = sLine.Replace(ControlChars.Tab, "    ") 'Replace tab with 4 spaces
        sLine = sLine.Trim

        Do While sLine.Length > 0
            sData = ""

            If regQuote.IsMatch(sLine) Then
                mc = regQuote.Matches(sLine)
                '
                ' "text",<rest of the line>
                '
                m = regQuote.Match(sLine)
                sData = m.Groups(2).Value
                sLine = m.Groups(5).Value
            ElseIf regQuoteLast.IsMatch(sLine) Then
                '
                ' "text"
                '
                m = regQuoteLast.Match(sLine)
                sData = m.Groups(2).Value
                sLine = ""
            ElseIf regNormal.IsMatch(sLine) Then
                '
                ' text,<rest of the line>
                '
                m = regNormal.Match(sLine)
                sData = m.Groups(1).Value
                sLine = m.Groups(3).Value
            ElseIf regNormalLast.IsMatch(sLine) Then
                '
                ' text
                '
                m = regNormalLast.Match(sLine)
                sData = m.Groups(0).Value
                sLine = ""
            Else
                '
                ' ERROR!!!!!
                '
                sData = ""
                sLine = ""
            End If

            sData = sData.Trim
            sLine = sLine.Trim

            If mData Is Nothing Then
                ReDim mData(0)
                idx = 0
            Else
                idx = mData.GetUpperBound(0) + 1
                ReDim Preserve mData(idx)
            End If

            mData(idx) = sData
        Loop
    End Sub
#End Region
#Region " Regular Expressions "
    '
    ' Set up Regular Expressions
    '
    Private Sub SetupRegEx()
        Dim sQuote As String = "^(%Q)(.*)(%Q)(\s*%S)(.*)$"
        Dim sNormal As String = "^([^%S]*)(\s*%S)(.*)$"
        Dim sQuoteLast As String = "^(%Q)(.*)(%Q$)"
        Dim sNormalLast As String = "^.*$"
        Dim sSep As String
        Dim sQual As String

        If Not regQuote Is Nothing Then regQuote = Nothing
        If Not regNormal Is Nothing Then regNormal = Nothing
        If Not regQuoteLast Is Nothing Then regQuoteLast = Nothing
        If Not regNormalLast Is Nothing Then regNormalLast = Nothing

        sSep = mSeparator
        sQual = mTextQualifier

        If InStr(".$^{[(|)]}*+?\", sSep) > 0 Then sSep = "\" & sSep
        If InStr(".$^{[(|)]}*+?\", sQual) > 0 Then sQual = "\" & sQual

        sQuote = sQuote.Replace("%S", sSep)
        sQuote = sQuote.Replace("%Q", sQual)
        sNormal = sNormal.Replace("%S", sSep)
        sQuoteLast = sQuoteLast.Replace("%Q", sQual)

        regQuote = New Regex(sQuote, RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
        regNormal = New Regex(sNormal, RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
        regQuoteLast = New Regex(sQuoteLast, RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
        regNormalLast = New Regex(sNormalLast, RegexOptions.IgnoreCase + RegexOptions.RightToLeft)
    End Sub
#End Region
#Region " Save As "
    '
    ' Save data as XML
    '
    Public Sub SaveAsXML(ByVal sXMLFile As String)
        If dsCSV Is Nothing Then Exit Sub
        dsCSV.WriteXml(sXMLFile)
    End Sub

    '
    ' Save data as CSV
    '
    Public Sub SaveAsCSV(ByVal sCSVFile As String)
        If dsCSV Is Nothing Then Exit Sub

        Dim dr As DataRow
        Dim sLine As String
        Dim sw As New StreamWriter(sCSVFile)
        Dim iCol As Integer

        For Each dr In dsCSV.Tables("CSVData").Rows
            sLine = ""
            For iCol = 0 To dsCSV.Tables("CSVData").Columns.Count - 1
                If sLine.Length > 0 Then sLine &= mSeparator
                If Not dr(iCol) Is DBNull.Value Then
                    If InStr(dr(iCol), mSeparator) > 0 Then
                        sLine &= mTextQualifier & dr(iCol) & mTextQualifier
                    Else
                        sLine &= dr(iCol)
                    End If
                End If
            Next

            sw.WriteLine(sLine)
        Next

        sw.Flush()
        sw.Close()
        sw = Nothing
    End Sub
#End Region
#Region " Properties "
    '
    ' Separator Property
    '
    Public Property Separator() As Char
        Get
            Return mSeparator
        End Get
        Set(ByVal Value As Char)
            mSeparator = Value
            SetupRegEx()
        End Set
    End Property

    '
    ' Qualifier Property
    '
    Public Property TextQualifier() As Char
        Get
            Return mTextQualifier
        End Get
        Set(ByVal Value As Char)
            mTextQualifier = Value
            SetupRegEx()
        End Set
    End Property

    '
    ' Dataset Property
    '
    Public ReadOnly Property CSVDataSet() As DataSet
        Get
            Return dsCSV
        End Get
    End Property
#End Region
#Region " Dispose and Finalize "
    '
    ' Dispose
    '
    Public Sub Dispose() Implements System.IDisposable.Dispose
        Dispose(True)
    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Disposed Then Exit Sub

        If disposing Then
            Disposed = True

            GC.SuppressFinalize(Me)
        End If

        If Not dsCSV Is Nothing Then
            dsCSV.Clear()
            dsCSV.Tables.Clear()
            dsCSV.Dispose()
            dsCSV = Nothing
        End If
    End Sub

    '
    ' Finalize
    '
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
#End Region
End Class
