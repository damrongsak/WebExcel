#Region ".NET Framework Class Import"
Imports System.Data
Imports System.Web
#End Region


Public Class DBUTIL

#Region "Internal member variables"
    Private _ConnectStr As String
    Private _DB_Provider As String
    Private _Params As New ArrayList
    Private _ConnectStr_Test1 As String = "Provider=MSDAORA;Data Source=TEST;User ID=ngvlogistic;Password=cngvlogistic"
    Private _ConnectStr_Test2 As String = "Provider=MSDAORA;Data Source=XE;User ID=tracking;Password=P@ssw0rd1"
    Private _isTest As String = "2" '1 test, 2 XE, 3...n default
#End Region

#Region "Declarations"

    Public Const opINSERT As Integer = 1
    Public Const opUPDATE As Integer = 2
    Public Const opDELETE As Integer = 3

    Public Enum FieldTypes
        ftNumeric = 1
        ftText = 2
        ftDate = 3
        ftDateTime = 6
        ftBinary = 7
    End Enum
#End Region

#Region "Properties"

    Public Property ConnectStr() As String
        Get
            Return _ConnectStr
        End Get
        Set(ByVal Value As String)
            _ConnectStr = Value
            Select Case _isTest
                Case "1"
                    _ConnectStr = _ConnectStr_Test1
                Case "2"
                    _ConnectStr = _ConnectStr_Test2
                Case Else
                    ' does some processing... 
                    Exit Select
            End Select
        End Set
    End Property

    Public Property DB_Provider() As String
        Get
            Return _DB_Provider
        End Get
        Set(ByVal Value As String)
            _DB_Provider = Value
        End Set
    End Property

#End Region

    Public Sub New()
        InitParams()
    End Sub

    Protected Overrides Sub Finalize()
        _Params.Clear()
        _Params = Nothing
        MyBase.Finalize()
    End Sub

    Public Function OpenConn(ByVal ConnectStr As String) As OleDb.OleDbConnection
        _ConnectStr = ConnectStr
        Select Case _isTest
            Case "1"
                _ConnectStr = _ConnectStr_Test1
            Case "2"
                _ConnectStr = _ConnectStr_Test2
            Case Else
                ' does some processing... 
                Exit Select
        End Select
        Return OpenDBConn()
    End Function

    Public Function OpenConn(ByVal DB_Provider As String, ByVal DB_DataSource As String, ByVal DB_UserName As String, ByVal DB_Password As String, ByVal DB_Name As String) As OleDb.OleDbConnection
        _DB_Provider = DB_Provider
        _ConnectStr = "Provider=" & DB_Provider & ";Data Source=" & DB_DataSource & ";User ID=" & DB_UserName & ";Password=" & DB_Password
        If DB_Name <> "" Then _ConnectStr += ";Initial Catalog=" & DB_Name
        Select Case _isTest
            Case "1"
                _ConnectStr = _ConnectStr_Test1
            Case "2"
                _ConnectStr = _ConnectStr_Test2
            Case Else
                ' does some processing... 
                Exit Select
        End Select
        Return OpenDBConn()
    End Function

    Private Function OpenDBConn() As OleDb.OleDbConnection
        Dim Conn As New OleDb.OleDbConnection
        Dim I As Integer

        Try
            'Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("th-TH")
            Select Case _isTest
                Case "1"
                    _ConnectStr = _ConnectStr_Test1
                Case "2"
                    _ConnectStr = _ConnectStr_Test2
                Case Else
                    ' does some processing... 
                    Exit Select
            End Select
            Conn.ConnectionString = _ConnectStr
            Conn.Open()

            If _DB_Provider = "" Then
                _DB_Provider = Conn.Provider.ToUpper
                I = _DB_Provider.IndexOf(".")
                If I >= 0 Then _DB_Provider = _DB_Provider.Substring(0, I)
            End If

            Return Conn
        Catch ex As Exception
            CloseConn(Conn)

            Throw ex
        End Try
    End Function

    '========================================
    ' Close Database Connection
    Public Sub CloseConn(ByRef Conn As OleDb.OleDbConnection)
        If Not (Conn Is Nothing) Then
            Try
                Conn.Close()
            Catch ex As Exception
            End Try
            Conn.Dispose()
            Conn = Nothing
        End If
    End Sub

    '========================================
    ' Begin Transaction
    Public Function BeginTrans(ByRef Conn As OleDb.OleDbConnection) As OleDb.OleDbTransaction
        If Not IsNothing(Conn) Then
            BeginTrans = Conn.BeginTransaction()
        Else
            Throw New Exception("Connection has not been initialized!")
        End If
    End Function

    '========================================
    ' Commit Transaction
    Public Sub CommitTrans(ByRef Trans As OleDb.OleDbTransaction)
        If Not IsNothing(Trans) Then
            Trans.Commit()
            Trans = Nothing
        End If
    End Sub

    '========================================
    ' Rollback Transaction
    Public Sub RollbackTrans(ByRef Trans As OleDb.OleDbTransaction)
        Try
            If Not IsNothing(Trans) Then
                Trans.Rollback()
            End If
        Catch
        End Try
        Trans = Nothing
    End Sub

    Public Function OpenDS(ByRef DS As DataSet, ByVal SQL As String, _
                    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataRow
        Dim DA As OleDb.OleDbDataAdapter = Nothing
        Try
            If DS Is Nothing Then
                DS = New DataSet
            Else
                DS.Clear()
            End If

            If Not IsNothing(Conn) Then
                DA = New OleDb.OleDbDataAdapter(SQL, Conn)
                If Not IsNothing(Trans) Then
                    DA.SelectCommand.Transaction = Trans
                End If
            Else
                DA = New OleDb.OleDbDataAdapter(SQL, _ConnectStr)
            End If
            DA.Fill(DS)
            If IsNothing(Conn) AndAlso Not IsNothing(DA.SelectCommand.Connection) Then
                CloseConn(DA.SelectCommand.Connection)
            End If
            DA.Dispose()
            DA = Nothing

            If (DS.Tables.Count > 0) AndAlso (DS.Tables(0).Rows.Count > 0) Then
                Return DS.Tables(0).Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            If Not IsNothing(DA) Then
                If IsNothing(Conn) AndAlso Not IsNothing(DA.SelectCommand.Connection) Then
                    CloseConn(DA.SelectCommand.Connection)
                End If
                DA.Dispose()
                DA = Nothing
            End If

            Throw (ex)
        End Try
    End Function


    '========================================
    ' Close DataSet
    Public Sub CloseDS(ByRef DS As DataSet)
        Try
            If Not IsNothing(DS) Then DS.Dispose()
        Catch
        End Try
        DS = Nothing
    End Sub

    Public Function OpenDT(ByRef DT As DataTable, ByVal SQL As String, _
                    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As DataRow
        Dim DA As OleDb.OleDbDataAdapter = Nothing

        Try

            If DT Is Nothing Then
                DT = New DataTable
            Else
                DT.Clear()
            End If

            If Not IsNothing(Conn) Then
                DA = New OleDb.OleDbDataAdapter(SQL, Conn)
                If Not IsNothing(Trans) Then
                    DA.SelectCommand.Transaction = Trans
                End If
            Else
                DA = New OleDb.OleDbDataAdapter(SQL, _ConnectStr)
            End If
            DA.Fill(DT)
            If IsNothing(Conn) AndAlso Not IsNothing(DA.SelectCommand.Connection) Then
                CloseConn(DA.SelectCommand.Connection)
            End If
            DA.Dispose()
            DA = Nothing

            If (DT.Rows.Count > 0) Then
                Return DT.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            If Not IsNothing(DA) Then
                If IsNothing(Conn) AndAlso Not IsNothing(DA.SelectCommand.Connection) Then
                    CloseConn(DA.SelectCommand.Connection)
                End If
                DA.Dispose()
                DA = Nothing
            End If

            Throw (ex)
        End Try
    End Function

    '========================================
    ' Close data table
    Public Sub CloseDT(ByRef DT As DataTable)
        Try
            If Not IsNothing(DT) Then DT.Dispose()
        Catch
        End Try
        DT = Nothing
    End Sub

    '========================================
    ' Execute an SQL Command
    Public Function ExecSQL(ByVal SQL As String, Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Integer
        Dim tmpConn As OleDb.OleDbConnection = Conn
        Dim cmd As OleDb.OleDbCommand = Nothing
        Dim rows As Integer = 0

        Try
            If IsNothing(Conn) Then tmpConn = OpenConn(_ConnectStr)
            cmd = New OleDb.OleDbCommand(SQL, tmpConn)
            If Not IsNothing(Trans) Then
                cmd.Transaction = Trans
            End If
            rows = cmd.ExecuteNonQuery()
            cmd.Dispose()
            cmd = Nothing
            If IsNothing(Conn) Then CloseConn(tmpConn)

            Return rows
        Catch ex As Exception
            If IsNothing(Conn) Then CloseConn(tmpConn)
            cmd.Dispose()
            cmd = Nothing
            Throw ex
        End Try
    End Function

    '========================================
    ' Lookup value from a SQL
    Public Function LookupSQL(ByVal SQL As String, _
                    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Object
        Dim DT As DataTable
        Dim DR As DataRow
        Dim Value As Object = ""

        DT = Nothing
        DR = OpenDT(DT, SQL, Conn, Trans)
        If Not IsNothing(DR) Then
            Value = DR.Item(0)
            If IsNothing(Value) OrElse IsDBNull(Value) Then Value = ""
        End If
        CloseDT(DT)

        Return Value
    End Function

    '========================================
    ' Lookup data in a table
    Public Function LookupData(ByVal usrTable As String, ByVal FieldName As String, ByVal usrCriteria As String, _
                    Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Object
        Dim SQL As String

        If usrCriteria <> "" Then
            SQL = " SELECT " & FieldName & "  FROM " & usrTable & " WHERE  " & usrCriteria
        Else
            SQL = " SELECT " & FieldName & "  FROM " & usrTable
        End If
        Return LookupSQL(SQL)
    End Function

    Public Function SQLDate(ByVal D As Object) As String
        Dim Y As Integer
        If IsDate(D) AndAlso (CDbl(D.ToOADate) > 0) Then
            Select Case _DB_Provider
                Case "MSDAORA"
                    Y = D.Year
                    If Y > 2500 Then Y = Y - 543
                    SQLDate = "TO_DATE('" & Y & "/" & D.Month & "/" & D.Day & "','YYYY/MM/DD')"
                Case Else
                    SQLDate = "convert(datetime," & (CDbl(D.ToOADate) - 2) & ")"
            End Select
        Else
            SQLDate = "NULL"
        End If
    End Function

    '========================================
    ' Format DateTime to Oracle SQL DateTime
    Public Function SQLDateTime(ByVal DT As Object) As String
        Dim Y As Integer
        If IsDate(DT) AndAlso (CDbl(DT.ToOADate) > 0) Then
            Select Case _DB_Provider
                Case "MSDAORA"
                    Y = DT.Year
                    If Y > 2500 Then Y = Y - 543
                    SQLDateTime = "TO_DATE('" & Y & "/" & DT.Month & "/" & DT.Day & " " & DT.Hour & ":" & DT.Minute & ":" & DT.Second & "','YYYY/MM/DD HH24:MI:SS')"
                Case Else
                    SQLDateTime = "convert(datetime," & (CDbl(DT.ToOADate) - 2) & ")"
                    'SQLDateTime = "'" & DT.ToString("yyyy-MM-dd HH:mm:ss") & "'"

            End Select
        Else
            SQLDateTime = "NULL"
        End If
    End Function

    '========================================
    ' Append a SQL criteria
    Public Function AddCriteria(ByRef CriteriaSQL As String, ByVal FieldName As String, ByVal FieldValue As Object, ByVal FieldType As FieldTypes) As Boolean
        Dim Oper As String = "="
        Dim FVal As String

        FVal = CStr(FieldValue)

        If FVal <> "" Then
            If FVal.IndexOf("%") >= 0 Then
                Oper = " LIKE "
                FieldType = FieldTypes.ftText
            End If
            If FVal.StartsWith("<") Then
                If FVal.Substring(1, 1) = ">" Then
                    Oper = "<>"
                    FieldValue = FVal.Substring(2)
                ElseIf FVal.Substring(1, 1) = "=" Then
                    Oper = "<="
                    FieldValue = FVal.Substring(2)
                Else
                    Oper = "<"
                    FieldValue = FVal.Substring(1)
                End If
            ElseIf FVal.StartsWith(">") Then
                If FVal.Substring(1, 1) = "=" Then
                    Oper = ">="
                    FieldValue = FVal.Substring(2)
                Else
                    Oper = ">"
                    FieldValue = FVal.Substring(1)
                End If
            End If

            Select Case FieldType
                Case FieldTypes.ftNumeric
                    If IsNumeric(FieldValue) Then
                        FVal = CStr(FieldValue)
                    ElseIf Oper = " LIKE " Then
                        FVal = "'" + Replace(CStr(FieldValue), "'", "''") + "'"
                    Else
                        FVal = ""
                    End If
                Case FieldTypes.ftText
                    If FVal <> "NULL" Then FVal = "'" + Replace(CStr(FieldValue), "'", "''") + "'"

                    FVal = "'" + Replace(CStr(FieldValue), "'", "''") + "'"
                    'FVal = "'%" + Replace(CStr(FieldValue), "'", "''") + "%'"
                    'Oper = " LIKE "
                Case FieldTypes.ftDate
                    If IsDate(FieldValue) AndAlso (CDbl(CDate(FieldValue).ToOADate) > 0) Then
                        FVal = SQLDate(CType(FieldValue, Date))
                    Else
                        FVal = ""
                    End If
                Case FieldTypes.ftDateTime
                    If IsDate(FieldValue) AndAlso (CDbl(CDate(FieldValue).ToOADate) > 0) Then
                        FVal = SQLDateTime(CType(FieldValue, Date))
                    Else
                        FVal = ""
                    End If
            End Select

            If FVal <> "" Then
                If CriteriaSQL <> "" Then CriteriaSQL += " AND "
                CriteriaSQL += FieldName + Oper + FVal
            End If
        End If
    End Function

    '========================================
    ' Append a SQL criteria
    Public Function AddCriteria2Condi(ByRef CriteriaSQL As String, ByVal FieldName As String, ByVal FieldName2 As String _
    , ByVal FieldValue As Object, ByVal FieldType As FieldTypes) As Boolean
        Dim Criteria As String = "", Criteria2 As String = ""

        AddCriteria(Criteria, FieldName, FieldValue, FieldType)
        AddCriteria(Criteria2, FieldName2, FieldValue, FieldType)
        If Criteria2 <> "" Then
            If CriteriaSQL <> "" Then CriteriaSQL &= " AND "
            CriteriaSQL &= "(" & Criteria & " OR " & Criteria2 & ") "
        ElseIf Criteria <> "" Then
            If CriteriaSQL <> "" Then CriteriaSQL &= " AND "
            CriteriaSQL &= Criteria
        End If
    End Function

    '========================================
    ' Append a SQL criteria Created By Aoy for search primary key are text
    Public Function AddCriteria2(ByRef CriteriaSQL As String, ByVal FieldName As String, ByVal FieldValue As Object, ByVal FieldType As FieldTypes) As Boolean
        Dim Oper As String = "="
        Dim FVal As String

        FVal = CStr(FieldValue)

        If FVal <> "" Then
            If FVal.IndexOf("%") >= 0 Then
                Oper = " LIKE "
                FieldType = FieldTypes.ftText
            End If
            If FVal.StartsWith("<") Then
                If FVal.Substring(1, 1) = ">" Then
                    Oper = "<>"
                    FieldValue = FVal.Substring(2)
                ElseIf FVal.Substring(1, 1) = "=" Then
                    Oper = "<="
                    FieldValue = FVal.Substring(2)
                Else
                    Oper = "<"
                    FieldValue = FVal.Substring(1)
                End If
            ElseIf FVal.StartsWith(">") Then
                If FVal.Substring(1, 1) = "=" Then
                    Oper = ">="
                    FieldValue = FVal.Substring(2)
                Else
                    Oper = ">"
                    FieldValue = FVal.Substring(1)
                End If
            End If

            Select Case FieldType
                Case FieldTypes.ftNumeric
                    If IsNumeric(FieldValue) Then
                        FVal = CStr(FieldValue)
                    End If
                Case FieldTypes.ftText
                    If FVal <> "NULL" Then FVal = "'" + Replace(CStr(FieldValue), "'", "''") + "'"

                    FVal = "'" + Replace(CStr(FieldValue), "'", "''") + "'"
                    'FVal = "'%" + Replace(CStr(FieldValue), "'", "''") + "%'"
                    'Oper = " LIKE "
                Case FieldTypes.ftDate
                    If IsDate(FieldValue) AndAlso (CDbl(CDate(FieldValue).ToOADate) > 0) Then
                        FVal = SQLDate(CType(FieldValue, Date))
                    End If
                Case FieldTypes.ftDateTime
                    If IsDate(FieldValue) AndAlso (CDbl(CDate(FieldValue).ToOADate) > 0) Then
                        FVal = SQLDateTime(CType(FieldValue, Date))
                    End If
            End Select

            If FVal <> "" Then
                If CriteriaSQL <> "" Then CriteriaSQL += " AND "
                CriteriaSQL += FieldName + Oper + FVal
            End If
        End If
    End Function

    '========================================
    Public Function AddCriteriaRange(ByRef CriteriaSQL As String, ByVal FieldName As String, ByVal FromValue As Object, ByVal ToValue As Object, ByVal FieldType As FieldTypes) As Boolean
        Dim FromVal As String = "", ToVal As String = ""
        Dim Oper1 As String = ""
        Dim Oper2 As String = ""
        If FromValue & "" <> "" Then
            FromVal = FromValue & ""
            If FromVal.StartsWith("<") Then
                If FromVal.Substring(1, 1) = ">" Then
                    Oper1 = "<>"
                    FromValue = FromVal.Substring(2)
                ElseIf FromVal.Substring(1, 1) = "=" Then
                    Oper1 = "<="
                    FromValue = FromVal.Substring(2)
                Else
                    Oper1 = "<"
                    FromValue = FromVal.Substring(1)
                End If
            ElseIf FromVal.StartsWith(">") Then
                If FromVal.Substring(1, 1) = "=" Then
                    Oper1 = ">="
                    FromValue = FromVal.Substring(2)
                Else
                    Oper1 = ">"
                    FromValue = FromVal.Substring(1)
                End If
            End If

            If ToValue & "" <> "" Then
                ToVal = ToValue & ""
                If ToVal.StartsWith("<") Then
                    If ToVal.Substring(1, 1) = ">" Then
                        Oper2 = "<>"
                        ToValue = ToVal.Substring(2)
                    ElseIf ToVal.Substring(1, 1) = "=" Then
                        Oper2 = "<="
                        ToValue = ToVal.Substring(2)
                    Else
                        Oper2 = "<"
                        ToValue = ToVal.Substring(1)
                    End If
                ElseIf ToVal.StartsWith(">") Then
                    If ToVal.Substring(1, 1) = "=" Then
                        Oper2 = ">="
                        ToValue = ToVal.Substring(2)
                    Else
                        Oper2 = ">"
                        ToValue = ToVal.Substring(1)
                    End If
                End If
            End If

            Select Case FieldType
                Case FieldTypes.ftNumeric
                    'If IsNumeric(FromValue) Then FromVal = CStr(FromValue)
                    'If IsNumeric(ToValue) Then ToVal = CStr(ToValue)
                    If IsNumeric(FromValue) Then
                        FromVal = CStr(FromValue)
                    Else
                        FromVal = ""
                    End If
                    If IsNumeric(ToValue) Then
                        ToVal = CStr(ToValue)
                    Else
                        ToVal = ""
                    End If
                    If IsNumeric(ToValue) Then ToVal = CStr(ToValue)
                Case FieldTypes.ftText
                    FromVal = "'" + Replace(CStr(FromValue), "'", "''") + "'"
                    If ToValue & "" <> "" Then ToVal = "'" + Replace(CStr(ToValue), "'", "''") + "'"
                    'If ToVal & "" <> "" Then ToVal = "'" + Replace(CStr(ToValue), "'", "''") + "'"
                Case FieldTypes.ftDate
                    'If IsDate(FromValue) AndAlso (CDbl(CDate(FromValue).ToOADate) > 0) Then FromVal = SQLDate(CType(FromValue, Date))
                    'If IsDate(ToValue) AndAlso (CDbl(CDate(ToValue).ToOADate) > 0) Then ToVal = SQLDate(DateAdd(DateInterval.Day, 1, CType(ToValue, Date)))
                    If IsDate(FromValue) AndAlso (CDbl(CDate(FromValue).ToOADate) > 0) Then
                        FromVal = SQLDate(CType(FromValue, Date))
                    Else
                        FromVal = ""
                    End If
                    If IsDate(ToValue) AndAlso (CDbl(CDate(ToValue).ToOADate) > 0) Then
                        ToVal = SQLDate(DateAdd(DateInterval.Day, 1, CType(ToValue, Date)))
                    Else
                        ToVal = ""
                    End If
                Case FieldTypes.ftDateTime
                    'If IsDate(FromValue) AndAlso (CDbl(CDate(FromValue).ToOADate) > 0) Then FromVal = SQLDateTime(CType(FromValue, Date))
                    ''If IsDate(ToValue) AndAlso (CDbl(CDate(ToValue).ToOADate) > 0) Then ToVal = SQLDateTime(CType(ToValue, Date))
                    'If IsDate(ToValue) AndAlso (CDbl(CDate(ToValue).ToOADate) > 0) Then ToVal = SQLDateTime(DateAdd(DateInterval.Day, 1, CType(ToValue, Date)))
                    If IsDate(FromValue) AndAlso (CDbl(CDate(FromValue).ToOADate) > 0) Then
                        FromVal = SQLDateTime(CType(FromValue, Date))
                    Else
                        FromVal = ""
                    End If

                    If IsDate(ToValue) AndAlso (CDbl(CDate(ToValue).ToOADate) > 0) Then
                        ToVal = SQLDateTime(DateAdd(DateInterval.Day, 1, CType(ToValue, Date)))
                    Else
                        ToVal = ""
                    End If
            End Select
        End If

        If FromVal & "" <> "" Then
            If ToVal & "" = "" Then
                If Oper1 <> "" Then
                    FromVal = Oper1 + FromVal
                Else
                    FromVal = FromValue
                End If
                Return AddCriteria(CriteriaSQL, FieldName, FromVal, FieldType)
            Else
                If CriteriaSQL <> "" Then CriteriaSQL += " AND "
                If FieldType = FieldTypes.ftDate Then
                    CriteriaSQL += "(" + FieldName + ">=" + FromVal + " AND " + FieldName + "<" + ToVal + ")"
                Else
                    If Oper1 <> "" AndAlso Oper2 <> "" Then
                        CriteriaSQL += "(" + FieldName + Oper1 + FromVal + " AND " + FieldName + Oper2 + ToVal + ")"
                    Else
                        CriteriaSQL += "(" + FieldName + " BETWEEN " + FromVal + " AND " + ToVal + ")"
                    End If
                End If
            End If
        End If
    End Function

    '========================================
    ' Format Value to SQL Command 
    Public Function SQLValue(ByVal Value As Object, ByVal DataType As Integer) As Object
        If Trim(Value & "") = "" OrElse Trim(Value & "") = "NULL" Then
            SQLValue = "NULL"
        Else
            Select Case DataType
                Case FieldTypes.ftDate
                    If UCase(TypeName(Value)) = "DATE" Then
                        SQLValue = SQLDate(Value)
                    ElseIf Value = "" Then
                        SQLValue = "NULL"
                    Else
                        SQLValue = UCase(Value & "")
                    End If
                Case FieldTypes.ftDateTime
                    If UCase(TypeName(Value)) Like "DATE*" Then
                        SQLValue = SQLDateTime(Value)
                    ElseIf Value = "" Then
                        SQLValue = "NULL"
                    Else
                        SQLValue = UCase(Value & "")
                    End If
                Case FieldTypes.ftNumeric
                    If IsNumeric(Value) Then
                        SQLValue = CDbl(Value)
                    ElseIf UCase(Value) Like "*.NEXTVAL" Then
                        SQLValue = Value
                    Else
                        SQLValue = "NULL"
                    End If
                Case FieldTypes.ftText
                    SQLValue = "'" & Replace(Value, "'", "''") & "'"
                Case Else
                    SQLValue = Value
            End Select
        End If
    End Function

    '========================================
    ' Add Parameter to INSERT/UPDATE SQL Command
    Public Sub AddSQL(ByVal operation As Integer, ByRef SQL1 As String, ByRef SQL2 As String, ByVal FieldName As String, ByVal FieldValue As Object, ByVal ColType As FieldTypes)
        Dim Data As String

        If FieldName <> "" Then
            Data = CStr(SQLValue(FieldValue, ColType))
            If operation = opINSERT Then
                If SQL1 <> "" Then
                    SQL1 = SQL1 + ", "
                    SQL2 = SQL2 + ", "
                End If
                SQL1 = SQL1 + FieldName
                SQL2 = SQL2 + Data
            Else    ' UPDATE
                If SQL1 <> "" Then SQL1 = SQL1 + ", "
                SQL1 = SQL1 + FieldName + "=" + Data
            End If
        End If
    End Sub

    Public Sub AddSQL2(ByVal operation As Integer, ByRef SQL1 As String, ByRef SQL2 As String, ByVal FieldName As String, ByVal FieldValue As Object, ByVal ColType As FieldTypes)
        'สำหรับ Mng - ถ้าไม่ส่งข้อมูลมา จะเป็น Nothing ซึ่งจะไม่ AddSQL
        If Not IsNothing(FieldValue) Then
            If FieldValue & "" = "NULL" Then
                AddSQL(operation, SQL1, SQL2, FieldName, FieldValue, ColType)
            Else
                Select Case ColType
                    Case FieldTypes.ftDate, FieldTypes.ftDateTime
                        If TypeName(FieldValue).ToLower() Like "date*" Then
                            AddSQL(operation, SQL1, SQL2, FieldName, FieldValue, ColType)
                        Else
                            AddSQL(operation, SQL1, SQL2, FieldName, AppDateValue(FieldValue), ColType)
                        End If
                    Case Else
                        AddSQL(operation, SQL1, SQL2, FieldName, FieldValue, ColType)
                End Select
            End If
            'Else
            '    'If ColType = FieldTypes.ftDate OrElse ColType = FieldTypes.ftDateTime Then
            '    '    FieldValue = "NULL"
            '    '    AddSQL(operation, SQL1, SQL2, FieldName, FieldValue, ColType)
            '    'End If
        End If
    End Sub

    '========================================
    ' Combine Parameter for INSERT/UPDATE SQL Command
    Public Function CombineSQL(ByVal operation As Integer, ByRef SQL1 As String, ByRef SQL2 As String, ByVal TableName As String, ByVal CriteriaSQL As String, Optional ByVal TimeStamp As Boolean = False) As String
        Dim SQL As String = ""
        Dim I As Integer

        'default Stamp updated record
        If TimeStamp Then
            AddSQL(operation, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME"), FieldTypes.ftText)
            AddSQL(operation, SQL1, SQL2, "DATE_UPDATED", Now, FieldTypes.ftDateTime)
        End If

        Select Case operation
            Case opINSERT
                SQL = "INSERT INTO " + TableName + " (" + SQL1 + ") VALUES (" + SQL2 + ")"
            Case opUPDATE
                SQL = "UPDATE " + TableName + " SET " + SQL1
                I = InStr(CriteriaSQL, "WHERE")
                If InStr(CriteriaSQL, "WHERE") > 0 Then
                    If UCase(Left(LTrim(CriteriaSQL), 3)) <> "AND" Then
                        SQL = SQL + " AND " + CriteriaSQL
                    Else
                        SQL = SQL + " " + CriteriaSQL
                    End If
                ElseIf CriteriaSQL <> "" Then
                    If UCase(Left(LTrim(CriteriaSQL), 3)) = "AND" Then
                        SQL = SQL + " WHERE " + Mid(LTrim(CriteriaSQL), 5)
                    Else
                        SQL = SQL + " WHERE " + CriteriaSQL
                    End If
                End If
            Case opDELETE
                SQL = "DELETE FROM " + TableName
                If CriteriaSQL <> "" Then
                    If UCase(Left(LTrim(CriteriaSQL), 3)) = "AND" Then
                        SQL = SQL + " WHERE " + Mid(LTrim(CriteriaSQL), 5)
                    Else
                        SQL = SQL + " WHERE " + CriteriaSQL
                    End If
                End If

        End Select
        If operation = opINSERT Then
        Else    ' UPDATE
        End If
        CombineSQL = SQL
    End Function

    '========================================
    ' Initialize stored procedure parameters
    Public Sub InitParams()
        _Params.Clear()
    End Sub

    '========================================
    ' Add a stored procedure parameter
    Public Sub AddParam(ByVal ParamName As String, ByVal Value As Object, ByVal DataType As FieldTypes)
        Dim P As New OleDb.OleDbParameter(ParamName, Value)

        Select Case DataType
            Case FieldTypes.ftNumeric
                P.OleDbType = OleDb.OleDbType.Numeric
            Case FieldTypes.ftText
                P.OleDbType = OleDb.OleDbType.VarChar
                P.DbType = DbType.String
            Case FieldTypes.ftDate
                If Not IsDate(Value) OrElse CDbl(CDate(Value).ToOADate) = 0 Then P.Value = Nothing
                P.OleDbType = OleDb.OleDbType.Date
            Case FieldTypes.ftDateTime
                If Not IsDate(Value) OrElse CDbl(CDate(Value).ToOADate) = 0 Then P.Value = Nothing
                P.OleDbType = OleDb.OleDbType.Date
        End Select

        _Params.Add(P)
    End Sub

    '========================================
    ' Execute a stored procedure
    Public Function ExecProc(ByVal ProcName As String, ByVal Conn As OleDb.OleDbConnection, _
                    Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Object
        Dim cmd As New OleDb.OleDbCommand(ProcName)
        Dim param As OleDb.OleDbParameter
        Dim I As Integer
        Dim results As Object = Nothing

        cmd.Connection = Conn
        cmd.CommandType = CommandType.StoredProcedure
        For I = 0 To _Params.Count - 1
            param = CType(_Params.Item(I), OleDb.OleDbParameter)
            cmd.Parameters.Add(param)
        Next

        cmd.ExecuteNonQuery()
        If cmd.Parameters.Contains("RETURN_VALUE") Then
            results = cmd.Parameters("RETURN_VALUE").Value
        End If

        cmd.Dispose()
        cmd = Nothing
        InitParams()

        Return results
    End Function

    '========================================
    ' Get Max Data

    Public Function GetMaxData(ByVal usrTable As String, ByVal FieldName As String, ByVal usrSQL As String) As Object
        Dim sql As String = ""

        sql = "SELECT MAX(" & FieldName & ") FROM " & usrTable
        If usrSQL & "" <> "" Then
            sql = sql & " WHERE " & usrSQL
        End If

        Return LookupSQL(sql)
    End Function

    Public Function ExecParamSQL(ByVal SQL As String, Optional ByVal Conn As OleDb.OleDbConnection = Nothing, _
                        Optional ByVal Trans As OleDb.OleDbTransaction = Nothing) As Integer
        Dim tmpConn As OleDb.OleDbConnection = Conn
        Dim cmd As OleDb.OleDbCommand = Nothing
        Dim param As OleDb.OleDbParameter
        Dim rows As Integer = 0
        Dim I As Integer

        Try
            If IsNothing(Conn) Then tmpConn = OpenConn(_ConnectStr)
            cmd = New OleDb.OleDbCommand(SQL, tmpConn)
            If Not IsNothing(Trans) Then
                cmd.Transaction = Trans
            End If

            For I = 0 To _Params.Count - 1
                param = CType(_Params.Item(I), OleDb.OleDbParameter)
                cmd.Parameters.Add(param)
            Next

            rows = cmd.ExecuteNonQuery()
            cmd.Dispose()
            cmd = Nothing
            If IsNothing(Conn) Then CloseConn(tmpConn)
            InitParams()

            Return rows

        Catch ex As Exception
            If IsNothing(Conn) Then CloseConn(tmpConn)
            cmd.Dispose()
            cmd = Nothing
            Throw ex
        End Try
    End Function

End Class


