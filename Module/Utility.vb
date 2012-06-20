#Region ".NET Framework Class Import"
'Imports System.Security
'Imports System.Security.Principal
'Imports System.Threading.Thread
Imports System.Data
Imports System.Web
Imports System.Drawing
Imports System.IO
#End Region


Public Module Utility

    Public Function GetCompleteMsg(ByVal MsgType As String, Optional ByVal UsrMsg As String = "", Optional ByVal Lang As String = "EN") As String
        Dim Msg As String = ""
        If UsrMsg = "" Then
            Select Case MsgType.ToUpper
                Case "LOAD"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการอ่านข้อมูล"
                    Else
                        Msg = "Successful in loading data."
                    End If
                Case "SAVE"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการบันทึกข้อมูล"
                    Else
                        Msg = "Successful in saving data."
                    End If
                Case "DELETE"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการลบข้อมูล"
                    Else
                        Msg = "Successful in deleting data."
                    End If
                Case "SUBMIT"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการยืนยันข้อมูล"
                    Else
                        Msg = "Successful in submitting data."
                    End If
                Case "COMMIT"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการมอบหมายข้อมูล"
                    Else
                        Msg = "Successful in committing data."
                    End If
                Case "CLOSE"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการปิดข้อมูล"
                    Else
                        Msg = "Successful in closing data."
                    End If
                Case "ADD"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการเพิ่มข้อมูล"
                    Else
                        Msg = "Successful in adding data."
                    End If
                Case "SEND"
                    If Lang = "TH" Then
                        Msg = "เสร็จสิ้นการส่งข้อมูล"
                    Else
                        Msg = "Successful in sending data."
                    End If
            End Select
        Else
            Msg = UsrMsg
        End If

        Msg = Msg.Replace("""", "'")
        Msg = Msg.Replace(vbCrLf, "\r\n")
        Msg = Msg.Replace(vbCr, "\r\n")
        Msg = Msg.Replace(vbLf, "\r\n")
        'WriteErrorLog(Msg)

        Return Msg
    End Function

    Public Function ValidateText(ByVal str As Object, Optional ByVal isReplaceAll As Boolean = True) As String
        Dim ret As String

        ret = str & ""
        ret = Replace(ret, ">", "")
        ret = Replace(ret, "<", "")
        ret = Replace(ret, "..", "")
        ret = Replace(ret, "--", "")
        ret = Replace(ret, "`", "")
        ret = Replace(ret, "'", "")
        ret = Replace(ret, "|", "")
        ret = Replace(ret, "&", "")
        ret = Replace(ret, ":", "")
        ret = Replace(ret, ";", "")
        ret = Replace(ret, "$", "")
        ret = Replace(ret, "@", "")
        ret = Replace(ret, ",", "")
        ret = Replace(ret, "\'", "")
        ret = Replace(ret, "\""", "")
        ret = Replace(ret, "+", "")
        ret = Replace(ret, "(", "")
        ret = Replace(ret, ")", "")
        ret = Replace(ret, """", "")
        ret = Replace(ret, "=", "")
        ret = Replace(ret, "<>", "")
        ret = Replace(ret, "()", "")
        ret = Replace(ret, "+", "")
        ret = Replace(ret, "#", "")
        If isReplaceAll Then
            ret = Replace(ret, "%", "")
            ret = Replace(ret, "*", "")
        End If

        Return ret & ""
    End Function

    Public Function GetErrorMsg(ByVal ex As Exception, Optional ByVal ProcName As String = "", Optional ByVal UsrMsg As String = "", Optional ByVal Lang As String = "EN") As String
        Dim ErrMsg As String = ""
        Dim Msg As String = ""
        Dim I, J As Integer

        If UsrMsg = "" Then
            If ProcName = "" Then
                J = ex.StackTrace.LastIndexOf("()")
                If J >= 0 Then
                    I = ex.StackTrace.Substring(0, J).LastIndexOf(".") + 1
                    ProcName = ex.StackTrace.Substring(I, J - I).ToUpper()
                End If
            Else
                ProcName = ProcName.ToUpper()
            End If

            If ProcName.StartsWith("LOAD") Then
                If Lang = "TH" Then
                    Msg = "เกิดข้อผิดพลาดระหว่างการอ่านข้อมูล"
                Else
                    Msg = "Error on loading data."
                End If
            ElseIf ProcName.StartsWith("SAVE") Then
                If Lang = "TH" Then
                    Msg = "เกิดข้อผิดพลาดระหว่างการบันทึกข้อมูล"
                Else
                    Msg = "Error on saving data."
                End If
            Else
                If Lang = "TH" Then
                    Select Case ProcName.ToUpper()
                        Case "DELETE"
                            Msg = "เกิดข้อผิดพลาดระหว่างการลบข้อมูล"
                        Case "OPENCONN"
                            Msg = "ไม่สามารถติดต่อฐานข้อมูลได้ กรุณาติดต่อผู้ดูแลระบบ"
                    End Select
                Else
                    Select Case ProcName.ToUpper()
                        Case "DELETE"
                            Msg = "Error on deleting data."
                        Case "OPENCONN"
                            Msg = "Unable to connect database.Please contact system administrator."
                    End Select
                End If

            End If
        Else
            Msg = UsrMsg
        End If

        If gDebugLevel & "" = "" Then ReadConfigurations()
        Select Case gDebugLevel
            Case "1"
                Msg += " => " & ex.Message
            Case "2"
                Msg += " => " & ex.ToString()
        End Select

        Msg = Msg.Replace("""", "'")
        Msg = Msg.Replace(vbCrLf, "\r\n")
        Msg = Msg.Replace(vbCr, "\r\n")
        Msg = Msg.Replace(vbLf, "\r\n")

        Try
            If InStr(Msg, "GoPage") <= 0 AndAlso InStr(Msg, "Thread was being aborted.") <= 0 Then
                'ErrMsg = Msg
                ErrMsg = ex.ToString()
                'ErrMsg = ErrMsg.Replace("""", "'")
                'ErrMsg = ErrMsg.Replace(vbCrLf, "\r\n")
                'ErrMsg = ErrMsg.Replace(vbCr, "\r\n")
                'ErrMsg = ErrMsg.Replace(vbLf, "\r\n")

                If UsrMsg <> "" Then ErrMsg &= vbCrLf & "[" & UsrMsg & "]"
                BLL.InsertAudit("Error", ErrMsg, HttpContext.Current.Session("USER_NAME") & "")
                'InsertAudit("Error", ErrMsg, HttpContext.Current.Session("USER_NAME") & "")
            End If
        Catch tex As Exception
        End Try

        Return Msg
    End Function

    Public Function NVL(ByVal Value As Object, Optional ByVal DefaultVal As Object = Nothing) As Object
        If IsDBNull(Value) OrElse IsNothing(Value) Then
            Return DefaultVal
        Else
            Return Value
        End If
    End Function

    Public Function GetDT(ByVal DS As DataSet) As DataTable
        If Not IsNothing(DS) AndAlso (DS.Tables.Count > 0) AndAlso (DS.Tables(0).Rows.Count > 0) Then
            Return DS.Tables(0)
        Else
            Return Nothing
        End If
    End Function
    'Aoy 9/12/51
    Public Function GetDT(ByVal Source As Object) As DataTable
        Dim DT As DataTable = Nothing
        Dim DType As String
        If IsNothing(Source) Then
            DT = Nothing
        Else
            DType = Source.GetType().Name
            Select Case DType
                Case "DataSet"
                    If (Source.Tables.Count > 0) AndAlso (Source.Tables(0).Rows.Count > 0) Then
                        DT = Source.Tables(0)
                    Else
                        DT = Nothing
                    End If
                Case "String"
                    Dim ds As New DataSet
                    If Not Source = "" Then
                        Dim stream As New StringReader(Source)
                        ds.ReadXml(stream)
                        DT = ds.Tables(0)
                    Else
                        DT = Nothing
                    End If
            End Select
        End If
        Return DT
    End Function

    Public Function GetDR(ByVal Source As Object) As DataRow
        Dim DR As DataRow = Nothing
        Dim DType As String
        If IsNothing(Source) Then
            DR = Nothing
        Else
            DType = Source.GetType().Name
            Select Case DType
                Case "DataSet"
                    If (Source.Tables.Count > 0) AndAlso (Source.Tables(0).Rows.Count > 0) Then
                        DR = Source.Tables(0).Rows(0)
                    Else
                        DR = Nothing
                    End If
                Case "DataTable"
                    If Not IsNothing(Source) AndAlso (Source.Rows.Count > 0) Then
                        DR = Source.Rows(0)
                    Else
                        DR = Nothing
                    End If
                Case "String"
                    If Not Source = "" Then
                        DR = GetDT(Source).Rows(0)
                    Else
                        DR = Nothing
                    End If
            End Select
        End If
        Return DR
    End Function

    Public Function GetDRV(ByVal DV As DataView) As DataRowView
        If Not IsNothing(DV) AndAlso (DV.Count > 0) Then
            Return DV.Item(0)
        Else
            Return Nothing
        End If
    End Function

    'Edit By Aoy 02/04/2552
    Public Function FormatDate(ByVal d As Object, ByVal fmt As String) As Object
        Dim DD, MM, YY As Integer

        If d & "" <> "" Then
            fmt = UCase(fmt)
            DD = Day(d)
            MM = Month(d)
            YY = Year(d)
            If YY > 2400 Then YY = YY - 543
            If InStr(1, fmt, "DD") > 0 Then
                If DD > 9 Then
                    fmt = Replace(fmt, "DD", DD)
                Else
                    fmt = Replace(fmt, "DD", "0" & DD)
                End If
            End If
            If InStr(1, fmt, "MM") > 0 Then
                If MM > 9 Then
                    fmt = Replace(fmt, "MM", MM)
                Else
                    fmt = Replace(fmt, "MM", "0" & MM)
                End If
            End If
            If InStr(1, fmt, "MONTH") > 0 Then fmt = Replace(fmt, "MONTH", EMonth(MM))
            If InStr(1, fmt, "MON") > 0 Then fmt = Replace(fmt, "MON", EMonth2(MM))
            If InStr(1, fmt, "YYYY") > 0 Then fmt = Replace(fmt, "YYYY", YY)
            If InStr(1, fmt, "YY") > 0 Then fmt = Replace(fmt, "YY", Right(YY, 2))
            If InStr(1, fmt, "HH") > 0 Then
                If Hour(d) > 9 Then
                    fmt = Replace(fmt, "HH", Hour(d))
                Else
                    fmt = Replace(fmt, "HH", "0" & Hour(d))
                End If
            End If

            If InStr(1, fmt, "MIN") > 0 Then
                If Minute(d) > 9 Then
                    fmt = Replace(fmt, "MIN", Minute(d))
                Else
                    fmt = Replace(fmt, "MIN", "0" & Minute(d))
                End If
            End If
            If InStr(1, fmt, "SS") > 0 Then
                If Second(d) > 9 Then
                    fmt = Replace(fmt, "SS", Second(d))
                Else
                    fmt = Replace(fmt, "SS", "0" & Second(d))
                End If
            End If
            ' แปลง format สำหรับวันที่ไทย
            If InStr(1, fmt, "วว") > 0 Then
                If DD > 9 Then
                    fmt = Replace(fmt, "วว", DD)
                Else
                    fmt = Replace(fmt, "วว", "0" & DD)
                End If
            End If
            If InStr(1, fmt, "ดดดด") > 0 Then
                fmt = Replace(fmt, "ดดดด", TMonth(MM))
            ElseIf InStr(1, fmt, "ดดด") > 0 Then
                fmt = Replace(fmt, "ดดด", TMonth2(MM))
            End If
            If InStr(1, fmt, "BBBB") > 0 Then
                fmt = Replace(fmt, "BBBB", YY + 543)
            ElseIf InStr(1, fmt, "BB") > 0 Then
                fmt = Replace(fmt, "BB", Right(YY + 543, 2))
            ElseIf InStr(1, fmt, "ปปปป") > 0 Then
                fmt = Replace(fmt, "ปปปป", YY + 543)
            ElseIf InStr(1, fmt, "ปป") > 0 Then
                fmt = Replace(fmt, "ปป", Right(YY + 543, 2))
            End If

            FormatDate = fmt
        Else
            FormatDate = ""
        End If
    End Function

    Function EMonth(ByVal m As Integer) As String
        Dim EM As String = ""
        Select Case m
            Case 1 : EM = "January"
            Case 2 : EM = "February"
            Case 3 : EM = "March"
            Case 4 : EM = "April"
            Case 5 : EM = "May"
            Case 6 : EM = "June"
            Case 7 : EM = "July"
            Case 8 : EM = "August"
            Case 9 : EM = "September"
            Case 10 : EM = "October"
            Case 11 : EM = "November"
            Case 12 : EM = "December"
        End Select
        EMonth = EM
        Return EMonth
    End Function

    Function EMonth2(ByVal m As Integer) As String
        Dim EM2 As String = ""
        Select Case m
            Case 1 : EM2 = "Jan"
            Case 2 : EM2 = "Feb"
            Case 3 : EM2 = "Mar"
            Case 4 : EM2 = "Apr"
            Case 5 : EM2 = "May"
            Case 6 : EM2 = "Jun"
            Case 7 : EM2 = "Jul"
            Case 8 : EM2 = "Aug"
            Case 9 : EM2 = "Sep"
            Case 10 : EM2 = "Oct"
            Case 11 : EM2 = "Nov"
            Case 12 : EM2 = "Dec"
        End Select
        Return EM2
    End Function

    Function TMonth(ByVal m As Integer) As String
        Dim ShMM As String = ""
        Select Case m
            Case 1 : ShMM = "มกราคม"
            Case 2 : ShMM = "กุมภาพันธ์"
            Case 3 : ShMM = "มีนาคม"
            Case 4 : ShMM = "เมษายน"
            Case 5 : ShMM = "พฤษภาคม"
            Case 6 : ShMM = "มิถุนายน"
            Case 7 : ShMM = "กรกฎาคม"
            Case 8 : ShMM = "สิงหาคม"
            Case 9 : ShMM = "กันยายน"
            Case 10 : ShMM = "ตุลาคม"
            Case 11 : ShMM = "พฤศจิกายน"
            Case 12 : ShMM = "ธันวาคม"
        End Select
        TMonth = ShMM
        Return TMonth
    End Function

    Function TMonth2(ByVal m As Integer) As String
        Dim ShMM2 As String = ""
        Select Case m
            Case 1 : ShMM2 = "ม.ค."
            Case 2 : ShMM2 = "ก.พ."
            Case 3 : ShMM2 = "มี.ค."
            Case 4 : ShMM2 = "เม.ย."
            Case 5 : ShMM2 = "พ.ค."
            Case 6 : ShMM2 = "มิ.ย."
            Case 7 : ShMM2 = "ก.ค."
            Case 8 : ShMM2 = "ส.ค."
            Case 9 : ShMM2 = "ก.ย."
            Case 10 : ShMM2 = "ต.ค."
            Case 11 : ShMM2 = "พ.ย."
            Case 12 : ShMM2 = "ธ.ค."
        End Select
        TMonth2 = ShMM2
        Return TMonth2
    End Function

    Public Function ShowEngMonth(ByVal MM As Object) As String
        ShowEngMonth = EMonth(MM)
    End Function

    Public Function ShowEngMM(ByVal MM As Object) As String
        ShowEngMM = EMonth2(MM)
    End Function

    Public Function ShowThaiMonth(ByVal MM As Object) As String
        ShowThaiMonth = TMonth(MM)
    End Function

    Public Function ShowThaiMM(ByVal MM As Object) As String
        ShowThaiMM = TMonth2(MM)
    End Function

    Public Function EngDateValue(ByVal S As Object) As Object
        Dim D, M, Y As Object
        Dim I, J, TS As Object
        Dim T As Date
        Dim Delim As String

        Delim = "/"
        S = Trim(S & "")
        I = InStr(1, S, Delim)
        If (Len(S) < 6) Or (I = 0) Then
            EngDateValue = Nothing
        Else
            D = CInt(Left(S, I - 1))
            J = InStr(I + 2, S, Delim)
            If J = 0 Then J = I
            If (I > 0) And (J - I + 1 > 0) Then
                M = Val(Mid(S, I + 1, J - I - 1))
            Else
                M = 0
            End If
            S = Mid(S, J + 1)
            If IsNumeric(S) Then
                Y = CInt(S)
            Else
                Y = 0
            End If
            I = InStr(1, S, " ")
            If S Like "*:*" Then
                TS = Trim(Mid(S, I + 1))
            Else
                TS = ""
            End If
            If Y > 2400 Then                ' i.e. 2548
                Y = Y - 543
            ElseIf Y > 30 And Y < 100 Then  ' i.e. 48
                Y = Y + 2500 - 543
            ElseIf Y <= 30 Then
                Y = Y + 2000
            End If

            If (D > 0) And (D < 32) And (M > 0) And (M < 13) And (Y > 1000) Then
                If Year(Now) > 2500 Then Y = Y + 543
                If TS <> "" Then
                    T = TimeValue(TS)
                    EngDateValue = New Date(Y, M, D, T.Hour, T.Minute, T.Second)
                Else
                    EngDateValue = DateSerial(Y, M, D)
                End If
            Else
                EngDateValue = Nothing
            End If
        End If
    End Function

    Public Function ThaiDateValue(ByVal S As Object) As Object
        Dim D, M, Y As Object
        Dim I, J, TS As Object
        Dim T As Date
        Dim Delim As String

        Delim = "/"
        If Not IsNothing(S) Then
            S = Trim(S & "")
            If S = "" Then
                ThaiDateValue = "NULL"
            Else
                I = InStr(1, S, Delim)
                If (Len(S) < 6) Or (I = 0) Then
                    ThaiDateValue = Nothing
                Else
                    D = CInt(Left(S, I - 1))
                    J = InStr(I + 2, S, Delim)
                    If J = 0 Then J = I
                    If (I > 0) And (J - I + 1 > 0) Then
                        '                M = TMonth2Num(Trim(Mid(S, I + 1, J - I - 1)))
                        M = Val(Mid(S, I + 1, J - I - 1))
                    Else
                        M = 0
                    End If
                    S = Mid(S, J + 1)
                    If IsNumeric(S) Then
                        Y = CInt(S)
                    Else
                        Y = 0
                    End If
                    I = InStr(1, S, " ")
                    If S Like "*:*" Then
                        TS = Trim(Mid(S, I + 1))
                    Else
                        TS = ""
                    End If

                    If Y = 0 And TS <> "" Then
                        S = S.replace(" " & TS, "")
                        If IsNumeric(S) Then
                            Y = CInt(S)
                        Else
                            Y = 0
                        End If
                    End If

                    If Y < 20 Then
                        Y = Y + 2543
                    ElseIf Y < 100 Then
                        Y = Y + 2500
                    ElseIf Y < 1900 Then
                        Y = 0
                    ElseIf Y < 2400 Then
                        Y = Y + 543
                    End If

                    If (D > 0) And (D < 32) And (M > 0) And (M < 13) And (Y > 2400) Then
                        If Year(Now) < 2500 Then Y = Y - 543
                        If TS <> "" Then
                            'Tai edit 26/04/2550
                            If Y > 2500 Then Y = Y - 543

                            T = TimeValue(TS)
                            ThaiDateValue = New Date(Y, M, D, T.Hour, T.Minute, T.Second)
                        Else
                            ThaiDateValue = DateSerial(Y, M, D)
                        End If
                    Else
                        ThaiDateValue = Nothing
                    End If
                End If
            End If
        Else
            ThaiDateValue = Nothing
        End If
    End Function

    Public Function ThaiDateTimeValue(ByVal S As Object, ByVal T As Object) As Object
        Dim D, M, Y As Integer
        Dim Hr, Min, Sec As Object
        Dim I, J As Integer
        Dim Delim As String

        Delim = "/"
        S = Trim(S & "")
        I = InStr(1, S, Delim)
        If (Len(S) < 6) Or (I = 0) Then
            ThaiDateTimeValue = Nothing
        Else
            D = CInt(Left(S, I - 1))
            J = InStr(I + 2, S, Delim)
            If J = 0 Then J = I
            If (I > 0) And (J - I + 1 > 0) Then
                M = Val(Mid(S, I + 1, J - I - 1))
            Else
                M = 0
            End If
            S = Mid(S, J + 1)
            If IsNumeric(S) Then
                Y = CInt(S)
            Else
                Y = 0
            End If
            I = InStr(1, S, " ")

            If Y < 20 Then
                Y = Y + 2543
            ElseIf Y < 100 Then
                Y = Y + 2500
            ElseIf Y < 1900 Then
                Y = 0
            ElseIf Y < 2400 Then
                Y = Y + 543
            End If

            If (D > 0) And (D < 32) And (M > 0) And (M < 13) And (Y > 2400) Then
                If Year(Now) < 2500 Then Y = Y - 543
                Hr = Left(T, InStr(T, ":") - 1)
                T = Right(T, Len(T) - Len(Hr) - 1)
                If InStr(T, ":") > 1 Then
                    Min = Left(T, InStr(T, ":") - 1)
                    T = Right(T, Len(T) - Len(Hr) - 1)
                    If Len(T) > 1 Then
                        Sec = T
                    Else
                        Sec = "0"
                    End If
                Else
                    Min = T
                    Sec = "0"
                End If

                ThaiDateTimeValue = DateSerial(Y, M, D) + " " + TimeSerial(Hr, Min, Sec)
                'ThaiDateTimeValue = DateSerial(Y - 543, M, D) + " " + TimeSerial(Hr, Min, Sec)
                'ThaiDateTimeValue = New Date(Y, M, D, Hr, Min, Sec)
            Else
                ThaiDateTimeValue = Nothing
            End If
        End If
    End Function

    Public Function ToNum(ByVal N As Object) As Double
        If IsNumeric(N) Then
            ToNum = CDbl(N)
        Else
            ToNum = 0
        End If
    End Function

    Public Function ToInt(ByVal N As Object) As Integer
        If IsNumeric(N) Then
            ToInt = CInt(N)
        Else
            ToInt = 0
        End If
    End Function

    Public Function ToLong(ByVal N As Object) As Long
        If IsNumeric(N) Then
            ToLong = CLng(N)
        Else
            ToLong = 0
        End If
    End Function

    Public Function ToDec(ByVal N As Object) As Decimal
        If IsNumeric(N) Then
            ToDec = CDec(N)
        Else
            ToDec = 0
        End If
    End Function

    Public Function SetYearID(ByVal YCell As Integer) As Object
        Dim YID As Object
        YID = Year(Today)
        If YID < 2400 Then YID += 543
        YID = Right(CStr(YID), YCell)
        SetYearID = YID
    End Function

    ' Manage Attach File
    Public Function GetFileType(ByVal FileName As String) As String
        If FileName & "" <> "" Then
            FileName = Mid(FileName, InStrRev(FileName, "\") + 1) & ""
            FileName = Mid(FileName, InStrRev(FileName, ".")) & ""
            Return FileName
        Else
            Return ""
        End If
    End Function

    Public Function GetFileName(ByVal FileName As String) As String
        If FileName & "" <> "" Then
            Return Mid(FileName, InStrRev(FileName, "\") + 1) & ""
        Else
            Return ""
        End If
    End Function

    Public Sub DeleteFile(ByVal FilePath As String)
        Dim FileDelete As System.IO.FileInfo = New System.IO.FileInfo(FilePath)
        If FileDelete.Exists Then FileDelete.Delete()
    End Sub

    Public Function AppDateValue(ByVal S As Object) As Object
        'กรณีแสดงผลเป็น 01-ม.ค.-2009 หรือ 01-MAR-2009 ให้ใช้ AppFormatSQLDate ก่อน
        AppDateValue = ThaiDateValue(S)
    End Function

    Public Function AppFormatDate(ByVal D As Object, Optional ByVal LangType As String = "EN") As String
        If LangType = "EN" Then
            AppFormatDate = FormatDate(D, "dd-MON-yyyy")
        Else
            AppFormatDate = FormatDate(D, "dd/mm/bbbb")
        End If
        'Dim errorMsg As String = String.Format(New CultureInfo("en-us", True), "There are some problems while trying to use the Cryptography Quick Start, please check the following error messages: " & Environment.NewLine & "{0}" & Environment.NewLine, e.Exception.Message)
    End Function

    Public Function AppFormatDateTime(ByVal D As Object, Optional ByVal LangType As String = "EN") As String
        If LangType = "EN" Then
            AppFormatDateTime = FormatDate(D, "dd-MON-yyyy HH:MIN:SS")
        Else
            AppFormatDateTime = FormatDate(D, "dd/mm/bbbb HH:MIN:SS")
        End If
        'Dim errorMsg As String = String.Format(New CultureInfo("en-us", True), "There are some problems while trying to use the Cryptography Quick Start, please check the following error messages: " & Environment.NewLine & "{0}" & Environment.NewLine, e.Exception.Message)
    End Function

    Public Function AppFormatTime(ByVal D As Object, Optional ByVal LangType As String = "EN") As String
        AppFormatTime = FormatDate(D, "HH:MIN")
        'Dim errorMsg As String = String.Format(New CultureInfo("en-us", True), "There are some problems while trying to use the Cryptography Quick Start, please check the following error messages: " & Environment.NewLine & "{0}" & Environment.NewLine, e.Exception.Message)
    End Function

    Public Function AppFormatSQLDate(ByVal D As Object, Optional ByVal LangType As String = "EN") As String
        If LangType = "EN" Then
            AppFormatSQLDate = FormatDate(D, "dd/mm/yyyy")
        Else
            AppFormatSQLDate = FormatDate(D, "dd/mm/bbbb")
        End If
        'Dim errorMsg As String = String.Format(New CultureInfo("en-us", True), "There are some problems while trying to use the Cryptography Quick Start, please check the following error messages: " & Environment.NewLine & "{0}" & Environment.NewLine, e.Exception.Message)
    End Function

    'Public Sub LoadMonthCombo(ByRef C As DropDownList, Optional ByVal nullvalue As Boolean = True, Optional ByVal DefaultMonth As String = "", Optional ByVal LangType As String = "TH")
    '    Dim I As Integer
    '    C.Items.Clear()
    '    If nullvalue Then
    '        C.Items.Insert(0, " ")
    '        C.Items(0).Value = ""
    '    End If
    '    If LangType & "" = "TH" Then
    '        For I = 1 To 12
    '            C.Items.Add(New ListItem(TMonth(I), I))
    '        Next
    '    Else
    '        For I = 1 To 12
    '            C.Items.Add(New ListItem(EMonth(I), I))
    '        Next
    '    End If

    '    If DefaultMonth <> "" Then
    '        C.SelectedValue = DefaultMonth
    '    End If
    'End Sub

    '-- AOR EDIT 13/12/2549 --
    '-- Tai Edit 27/04/2550 --
    Public Sub LoadMonthCombo(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal DefaultMonth As String = "", Optional ByVal LangType As String = "TH", Optional ByVal BlankText As String = "")
        Dim I As Integer

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If
        If LangType & "" = "TH" Then
            For I = 1 To 12
                C.Items.Add(New ListItem(TMonth(I), I))
            Next
        Else
            For I = 1 To 12
                C.Items.Add(New ListItem(EMonth(I), I))
            Next
        End If

        If DefaultMonth <> "" Then
            C.SelectedValue = DefaultMonth
        Else
            C.SelectedIndex = -1
        End If
    End Sub

    'Public Sub Load2MonthCombo(ByRef C As DropDownList, Optional ByVal nullvalue As Boolean = True, Optional ByVal DefaultMonth As Object = Nothing, Optional ByVal LangType As String = "TH")
    '    Dim I, MM As Integer

    '    C.Items.Clear()
    '    If nullvalue Then
    '        C.Items.Insert(0, " ")
    '        C.Items(0).Value = ""
    '    End If
    '    If LangType & "" = "TH" Then
    '        For I = 1 To 12
    '            C.Items.Add(New ListItem(TMonth(I), Format(I, "00")))
    '        Next
    '    Else
    '        For I = 1 To 12
    '            C.Items.Add(New ListItem(EMonth(I), Format(I, "00")))
    '        Next
    '    End If

    '    If DefaultMonth & "" <> "" Then
    '        C.SelectedValue = Format(DefaultMonth, "00") & ""
    '    End If
    'End Sub

    Public Sub LoadYearCombo(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal DefaultYear As String = "", Optional ByVal LangType As String = "EN", Optional ByVal BlankText As String = "", Optional ByVal num As Integer = 10, Optional ByVal BeforeNum As Integer = 0)
        Dim i As Integer
        Dim cnt As Integer = 0
        Dim sYear, bYear, eYear As Integer

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If
        sYear = System.DateTime.Now.Year
        If DefaultYear <> "" AndAlso IsNumeric(DefaultYear) Then
            If LangType = "TH" Then
                If CInt(DefaultYear) < 2500 Then DefaultYear = (CInt(DefaultYear) + 543) & ""
            Else
                If CInt(DefaultYear) > 2500 Then DefaultYear = (CInt(DefaultYear) - 543) & ""
            End If
        End If
        If LangType = "TH" Then
            If sYear < 2500 Then sYear += 543
        Else
            If sYear > 2500 Then sYear -= 543
        End If
        bYear = sYear - num
        eYear = sYear + BeforeNum

        i = eYear
        'Do While i >= bYear
        '    C.Items.Add(New ListItem(i, i))
        '    i -= 1
        'Loop
        'i = eYear
        Do While i >= bYear
            C.Items.Add(New ListItem(i, i))
            i = i - 1
        Loop
        C.SelectedValue = DefaultYear
    End Sub

    Public Sub InitYear(ByRef txtYear As TextBox, Optional ByVal nullvalue As Boolean = False)
        If Now.Year > 2500 Then
            txtYear.Text = Now.Year
        Else
            txtYear.Text = Now.Year + 543
        End If
    End Sub

    Public Sub InitMonthYear(ByRef lstMonth As DropDownList, ByRef txtYear As TextBox, Optional ByVal nullvalue As Boolean = False)
        Dim RefDate As Object
        RefDate = Now.Date.AddDays(-10)
        LoadMonthCombo(lstMonth, nullvalue, Month(RefDate))
        If RefDate.Year > 2500 Then
            txtYear.Text = RefDate.Year
        Else
            txtYear.Text = RefDate.Year + 543
        End If
    End Sub

    Public Sub LoadMMCombo(ByRef C As Object, Optional ByVal nullvalue As Boolean = True, Optional ByVal DefaultMonth As String = "", Optional ByVal LangType As String = "TH")
        Dim I As Integer

        C.Items.Clear()
        If nullvalue Then
            C.Items.Insert(0, " ")
            C.Items(0).Value = ""
        End If
        If LangType & "" = "TH" Then
            For I = 1 To 12
                C.Items.Add(New ListItem(TMonth(I), I))
            Next
        Else
            For I = 1 To 12
                C.Items.Add(New ListItem(EMonth(I), I))
            Next
        End If

        If DefaultMonth <> "" Then
            C.Value = DefaultMonth
        End If
    End Sub

    Public Sub LoadMMCombo2(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal DefaultMonth As String = "" _
    , Optional ByVal BlankText As String = "")

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If

        For I As Integer = 1 To 12
            C.Items.Add(New ListItem(Right("0" & I.ToString, 2), I))
        Next
        C.SelectedValue = DefaultMonth
    End Sub

    Public Sub LoadDDCombo(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal DefaultDay As String = "" _
    , Optional ByVal BlankText As String = "")

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If

        For I As Integer = 1 To 31
            C.Items.Add(New ListItem(Right("0" & I.ToString, 2), I))
        Next
        C.SelectedValue = DefaultDay
    End Sub

    Public Sub InitMMYear(ByRef lstMonth As Object, ByRef txtYear As Object, Optional ByVal nullvalue As Boolean = False)
        Dim RefDate As Object
        RefDate = Now.Date.AddDays(-10)
        LoadMMCombo(lstMonth, nullvalue, Month(RefDate))
        If RefDate.Year > 2500 Then
            txtYear.Value = RefDate.Year
        Else
            txtYear.Value = RefDate.Year + 543
        End If
    End Sub

    'Public Sub Init2MonthYear(ByRef lstMonth As DropDownList, ByRef txtYear As TextBox, Optional ByVal nullvalue As Boolean = False)
    '    Dim RefDate
    '    RefDate = Now.Date.AddDays(-10)
    '    Load2MonthCombo(lstMonth, nullvalue, Month(RefDate))
    '    If RefDate.Year > 2500 Then
    '        txtYear.Text = RefDate.Year
    '    Else
    '        txtYear.Text = RefDate.Year + 543
    '    End If
    'End Sub

    ' **********************************************
    Public Function AppDateToSapDate(ByVal AppDate As Object) As String
        Dim SapDate As String
        Dim D As Date

        If AppDate & "" = "" Then
            SapDate = ""
        Else
            Try
                D = AppDateValue(AppDate)
                SapDate = FormatDate(D, "yyyymmdd")
            Catch ex As Exception
                SapDate = "00000000"
            End Try
        End If

        Return SapDate
    End Function

    Public Function SapDateToAppDate(ByVal SapDate As Object) As String
        Dim DD, MM, YY As String
        Dim AppDate As String

        If SapDate & "" = "" Then
            AppDate = ""
        Else
            If SapDate Like "??.??.????" Then
                DD = SapDate.Substring(0, 2)
                MM = SapDate.Substring(3, 2)
                YY = SapDate.Substring(6, 4)
            Else
                DD = SapDate.Substring(6, 2)
                MM = SapDate.Substring(4, 2)
                YY = SapDate.Substring(0, 4)
            End If
            If (DD = "00") Then
                AppDate = ""
            Else
                AppDate = DD & "/" & MM & "/" & (Val(YY) + 543)
            End If
        End If

        Return AppDate
    End Function

    Public Function CSapDate(ByVal DateVal As Object) As String
        Dim SapDate As String

        If DateVal & "" = "" Then
            SapDate = "00000000"
        Else
            Try
                SapDate = FormatDate(DateVal, "yyyymmdd")
            Catch ex As Exception
                SapDate = "00000000"
            End Try
        End If

        Return SapDate
    End Function

    ' **********************************************
    Public Sub ClearObject(ByRef obj As Object)
        Try
            obj.Dispose()
        Catch ex As Exception
        End Try
        obj = Nothing
    End Sub

    Public Sub LockCtrl(ByVal C As Object)
        Dim ObjName As String
        If Not IsNothing(C.ID) Then
            ObjName = C.ID.Substring(0, 3).ToLower()
            Select Case ObjName
                Case "txt"
                    If TypeOf (C) Is TextBox Then
                        C.ReadOnly = True
                        C.CssClass = "txtReadOnly"
                        'C.BackColor = Color.FromName("#FFFFC0")
                        'If C.TextMode = TextBoxMode.MultiLine Then
                        '    C.CssClass = "txtReadOnly"
                        'End If
                    Else
                        C.Disabled = True
                    End If
                Case "lst", "cbo", "ddl"
                    C.Enabled = False
                    'C.BackColor = Color.FromName("#FFFFC0")
                Case "rbl"
                    C.Enabled = False
                Case "rdo", "rdb"
                    If TypeOf (C) Is HtmlInputRadioButton Then
                        C.Disabled = True
                    Else
                        C.Enabled = False
                    End If
                Case "cbl"
                    C.Enabled = False
                Case "btn"
                    C.Visible = False
                Case "img"
                    C.Visible = False
                Case "chk"
                    If TypeOf (C) Is CheckBox Then
                        C.ReadOnly = True
                        'C.BackColor = Color.FromName("#FFFFC0")
                    Else
                        C.Disabled = True
                    End If
            End Select
        End If
    End Sub

    Public Sub LockControls(ByVal Controls As ControlCollection)
        Dim C As Object

        For Each C In Controls
            LockCtrl(C)
        Next
    End Sub

    Public Function CRDate(ByVal D As Object) As String
        CRDate = FormatDate(D, "yyyy/mm/dd")
    End Function

    Public Function ValidateData(ByVal str As Object, Optional ByVal isReplaceAll As Boolean = True _
    , Optional ByVal CommaFlag As Boolean = True, Optional ByVal IsEmail As Boolean = True _
   , Optional ByVal SeperateFlag As Boolean = True, Optional ByVal IsTime As Boolean = False _
   , Optional ByVal IsCriteria As Boolean = False, Optional ByVal StarFlag As Boolean = False) As String
        Dim ret As String

        ret = str & ""
        ret = Replace(ret, "..", "")
        ret = Replace(ret, "--", "")
        ret = Replace(ret, "`", "")
        ret = Replace(ret, "&", "")
        If Not IsTime Then
            ret = Replace(ret, ":", "")
        End If
        If SeperateFlag Then
            ret = Replace(ret, "|", "")
            ret = Replace(ret, ";", "")
        End If
        ret = Replace(ret, "$", "")
        If Not IsEmail Then
            ret = Replace(ret, "@", "")
        End If
        ret = Replace(ret, "\'", "")
        ret = Replace(ret, "\""", "")
        ret = Replace(ret, "+", "")
        ret = Replace(ret, "<CR>", "")
        ret = Replace(ret, "<LF>", "")
        ret = Replace(ret, "()", "")
        ret = Replace(ret, "+", "")
        ret = Replace(ret, "#", "")
        If CommaFlag = True Then
            ret = Replace(ret, ",", "")
        End If

        If isReplaceAll Then
            ret = Replace(ret, "%", "")
            If Not StarFlag Then ret = Replace(ret, "*", "")
        End If

        If Not IsCriteria Then
            ret = Replace(ret, """", "")
            ret = Replace(ret, "'", "")
            ret = Replace(ret, "(", "")
            ret = Replace(ret, ")", "")
            ret = Replace(ret, "<>", "")
            ret = Replace(ret, "=", "")
            ret = Replace(ret, ">", "")
            ret = Replace(ret, "<", "")
        End If

        Return ret & ""
    End Function

    Public Function FormatSearchText(ByVal searchText As String, Optional ByVal searchBeginning As Boolean = False) As String
        Dim RET As String = ""

        If searchText <> "" Then
            '15/09/2551
            searchText = ValidateData(searchText, False)
            searchText = Replace(searchText, "*", "%")
            If Not (InStr(searchText, "%") > 0) Then
                If searchBeginning Then
                    searchText = searchText & "%"
                Else
                    searchText = "%" & searchText & "%"
                End If
            End If

            RET = searchText
        End If

        Return RET
    End Function

    Public Function FormatSaveData(ByVal Data As String, Optional ByVal DefaultValue As String = "NULL") As String
        Dim RET As String = ""

        If Data = "" Then
            RET = DefaultValue
        Else
            RET = Data
        End If

        Return RET
    End Function

    Public Sub UnLockControls(ByVal Controls As ControlCollection)
        Dim C As Object

        For Each C In Controls
            UnLockCtrl(C)
        Next
    End Sub

    Public Sub UnLockCtrl(ByVal C As Object)
        Dim ObjName As String
        If Not IsNothing(C.ID) Then
            ObjName = C.ID.Substring(0, 3).ToLower
            Select Case ObjName
                Case "txt"
                    If TypeOf (C) Is TextBox Then
                        C.ReadOnly = False
                        C.CssClass = ""
                    Else
                        C.Disabled = False
                    End If
                Case "lst", "cbo", "ddl"
                    C.Enabled = True
                    C.BackColor = Color.FromName("#FFFFFF")
                Case "rbl"
                    C.Enabled = True
                Case "rdo", "rdb"
                    C.Enabled = True
                Case "cbl"
                    C.Enabled = True
                Case "btn"
                    C.Visible = True
                Case "img"
                    C.Visible = True
                Case "chk"
                    If TypeOf (C) Is CheckBox Then
                        C.ReadOnly = False
                        C.BackColor = Color.FromName("#FFFFFF")
                    Else
                        C.Disabled = False
                    End If
            End Select
        End If
    End Sub

    Public Function AppFormatDateLong(ByVal D As Object, Optional ByVal LangType As String = "TH") As String
        If LangType = "EN" Then
            AppFormatDateLong = FormatDate(D, "dd MON yyyy")
        Else
            AppFormatDateLong = FormatDate(D, "dd ดดดด bbbb")
        End If
        'Dim errorMsg As String = String.Format(New CultureInfo("en-us", True), "There are some problems while trying to use the Cryptography Quick Start, please check the following error messages: " & Environment.NewLine & "{0}" & Environment.NewLine, e.Exception.Message)
    End Function

    Public Sub LoadList(ByRef C As Object, ByVal Data As Object, ByVal DescField As String, ByVal ValueField As String, Optional ByVal IncBlank As Boolean = False, Optional ByVal IncDesc As String = "", Optional ByVal IncValue As Object = Nothing, Optional ByVal IsLastRecord As Boolean = False, _
 Optional ByVal IncTotal As Boolean = False, Optional ByVal TotalDesc As String = "", Optional ByVal TotalValue As Object = Nothing, Optional ByVal IncTotal1 As Boolean = False, Optional ByVal TotalDesc1 As String = "", Optional ByVal TotalValue1 As Object = Nothing)
        Dim DT As DataTable
        Dim DR As DataRow

        Try
            If TypeOf (Data) Is DataSet Then
                DT = GetDT(Data)
            Else
                DT = Data
            End If

            If IncBlank Then
                DR = DT.NewRow
                DR(DescField) = IIf(IncDesc = "", DBNull.Value, IncDesc)
                If Not IsNothing(IncValue) Then DR(ValueField) = IncValue
                If IsLastRecord = False Then
                    DT.Rows.InsertAt(DR, 0)
                Else
                    DT.Rows.Add(DR)
                End If
            End If

            If IncTotal Then
                DR = DT.NewRow
                DR(DescField) = IIf(TotalDesc = "", DBNull.Value, TotalDesc)
                If Not IsNothing(TotalValue) Then DR(ValueField) = TotalValue
                DT.Rows.Add(DR)
            End If

            If IncTotal1 Then
                DR = DT.NewRow
                DR(DescField) = IIf(TotalDesc1 = "", DBNull.Value, TotalDesc1)
                If Not IsNothing(TotalValue1) Then DR(ValueField) = TotalValue1
                DT.Rows.Add(DR)
            End If

            If TypeOf C Is HtmlSelect Then
                For Each DR In DT.Rows
                    C.Items.Add(New ListItem(DR(DescField) & "", DR(ValueField) & ""))
                Next
            Else
                C.Items.Clear()
                C.DataSource = DT
                C.DataTextField = DescField
                C.DataValueField = ValueField
                C.DataBind()
            End If

            ClearObject(Data)
            ClearObject(DT)

        Catch ex As Exception
            If Not IsNothing(C) Then
                C.Items.Clear()
            End If
        End Try
    End Sub

    Public Sub LoadListSQL(ByRef C As Object, ByVal SQL As String, ByVal DescField As String, ByVal ValueField As String, Optional ByVal IncBlank As Boolean = False, Optional ByVal IncDesc As String = "", Optional ByVal IncValue As Object = Nothing, Optional ByVal IsLastRecord As Boolean = False, _
                        Optional ByVal IncTotal As Boolean = False, Optional ByVal TotalDesc As String = "", Optional ByVal TotalValue As Object = Nothing, Optional ByVal IncTotal1 As Boolean = False, Optional ByVal TotalDesc1 As String = "", Optional ByVal TotalValue1 As Object = Nothing)
        Dim DT As DataTable = Nothing
        Try
            DT = DAL.QueryData(SQL)
            LoadList(C, DT, DescField, ValueField, IncBlank, IncDesc, IncValue, IsLastRecord, IncTotal, TotalDesc, TotalValue, IncTotal1, TotalDesc1, TotalValue1)
        Catch
        End Try
        ClearObject(DT)
    End Sub

    Public Sub LoadListLevel(ByRef C As Object, ByVal Data As Object, ByVal DescField As String, ByVal ValueField As String, Optional ByVal LevelField As String = "", Optional ByVal IncBlank As Boolean = False, Optional ByVal IncValue As String = "")
        Dim DT As DataTable
        Dim DR As DataRow
        Dim strValue, strDesc As String
        Dim i As Integer
        Try
            If TypeOf (Data) Is DataSet Then
                DT = GetDT(Data)
            Else
                DT = Data
            End If

            If IncBlank Then
                DR = DT.NewRow
                DR(DescField) = IIf(IncValue = "", DBNull.Value, IncValue)
                DT.Rows.InsertAt(DR, 0)
            End If
            C.Items.Clear()

            If LevelField <> "" Then
                For Each DR In DT.Rows
                    'If DR(LevelField) & "" = "1" Then
                    '    strDesc = IIf(IsBold1Level, "<b>" & DR(DescField) & "</b>", DR(DescField) & "")
                    'Else
                    '    strDesc = StrDup(CInt(DR(LevelField) & ""), "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") & DR(DescField) & ""
                    'End If
                    i = CInt(DR(LevelField) & "")
                    If i > 1 Then
                        strDesc = StrDup(i, "-") & " " & DR(DescField) & ""
                    Else
                        strDesc = DR(DescField) & ""
                    End If
                    C.Items.Add(New ListItem(strDesc, DR(ValueField) & ""))
                Next
            Else
                For Each DR In DT.Rows
                    strValue = DR(ValueField) & ""
                    strDesc = DR(DescField) & ""
                    If ValueField = "UNITCODE" Then
                        If strValue <> "" Then
                            i = Len(strValue)
                            Do While (i > 1) And (Mid(strValue, i, 1) = "0")
                                i = i - 1
                            Loop
                            If i > 1 Then
                                strDesc = StrDup((i - 1), "-") + "> " & strDesc
                            End If
                        End If
                    End If
                    C.Items.Add(New ListItem(strDesc, strValue))
                Next
            End If

        Catch
            If Not IsNothing(C) Then
                C.Items.Clear()
            End If
        End Try
        'ClearObject(Data)
    End Sub

    Public Sub SetListValue(ByRef C As ListControl, ByVal SelectedValue As Object)
        Try
            C.SelectedValue = SelectedValue
        Catch ex As Exception
            C.SelectedIndex = -1
        End Try
    End Sub

    Public Sub SetCaseDropDownValue(ByRef C As AjaxControlToolkit.CascadingDropDown, ByVal SelectedValue As Object)
        Try
            C.SelectedValue = SelectedValue
        Catch ex As Exception
            C.SelectedValue = ""
        End Try
    End Sub

    Public Sub GetListValue(ByRef C As ListControl, ByRef Value As Object, Optional ByVal DefaultVal As String = "")
        Try
            Value = C.SelectedValue
        Catch ex As Exception
            Value = DefaultVal
        End Try
    End Sub

    Public Sub GetCaseDropDownValue(ByRef C As AjaxControlToolkit.CascadingDropDown, ByRef Value As Object, Optional ByVal DefaultVal As String = "")
        Try
            Value = C.SelectedValue
        Catch ex As Exception
            Value = DefaultVal
        End Try
    End Sub


    Public Sub BindDGData(ByRef DG As DataGrid, ByRef dataSource As Object, Optional ByVal AddBlankRow As Boolean = False)
        Dim DS As DataSet
        Dim DT As DataTable
        Dim DV As DataView

        Try
            If Not IsNothing(dataSource) Then
                Select Case TypeName(dataSource).ToUpper
                    Case "DATASET"
                        DS = CType(dataSource, DataSet)
                        If DS.Tables.Count > 0 Then
                            DT = DS.Tables(0)
                            If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                                DT.Rows.Add(DT.NewRow)     ' Add blank row
                            End If
                            DG.DataSource = DT
                            DG.DataBind()
                        End If
                    Case "DATATABLE"
                        DT = CType(dataSource, DataTable)
                        If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                            DT.Rows.Add(DT.NewRow)     ' Add blank row
                        End If
                        DG.DataSource = DT
                        DG.DataBind()
                    Case "DATAVIEW"
                        DV = CType(dataSource, DataView)
                        If AddBlankRow AndAlso DV.Count = 0 Then
                            DV.AddNew()     ' Add blank row
                        End If
                        DG.DataSource = DV
                        DG.DataBind()
                End Select
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub BindGVData(ByRef GV As GridView, ByRef dataSource As Object, Optional ByVal AddBlankRow As Boolean = False)
        Dim DS As DataSet
        Dim DT As DataTable
        Dim DV As DataView

        Try
            If Not IsNothing(GV) AndAlso Not IsNothing(dataSource) Then
                Select Case TypeName(dataSource).ToUpper()
                    Case "DATASET"
                        DS = CType(dataSource, DataSet)
                        If DS.Tables.Count > 0 Then
                            DT = DS.Tables(0)
                            If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                                DT.Rows.Add(DT.NewRow)     ' Add blank row
                            End If
                            GV.DataSource = DT
                            GV.DataBind()
                        End If
                    Case "DATATABLE"
                        DT = CType(dataSource, DataTable)
                        If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                            DT.Rows.Add(DT.NewRow)     ' Add blank row
                        End If
                        GV.DataSource = DT
                        GV.DataBind()
                    Case "DATAVIEW"
                        DV = CType(dataSource, DataView)
                        If AddBlankRow AndAlso DV.Count = 0 Then
                            DV.AddNew()     ' Add blank row
                        End If
                        GV.DataSource = DV
                        GV.DataBind()
                End Select
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function DataEncrypt(ByVal Data As String) As String
        Dim Encoder As New System.Text.ASCIIEncoding
        Dim ret As String

        Try
            ret = Convert.ToBase64String(Encoder.GetBytes(Data.ToCharArray()))
        Catch ex As Exception
            ret = ""
        Finally
            ClearObject(Encoder)
        End Try

        Return ret
    End Function

    Public Function DataDecrypt(ByVal Data As String) As String
        Dim Encoder As New System.Text.ASCIIEncoding
        Dim ret As String

        Try
            ret = Encoder.GetString(Convert.FromBase64String(Data))
        Catch ex As Exception
            ret = ""
        Finally
            ClearObject(Encoder)
        End Try

        Return ret
    End Function

    Public Sub GetStartEndDate(ByVal year As Integer, ByVal month As Integer, ByRef sDate As Date, ByRef eDate As Date)
        Try
            If year > 2500 Then year -= 543

            ClearObject(sDate)
            sDate = New Date(year, month, 1)

            ClearObject(eDate)
            eDate = New Date(year, month, Date.DaysInMonth(year, month))
        Catch ex As Exception
        End Try
    End Sub

    Public Function WriteTextFile(ByVal Filename As String, ByVal TextData As Object, Optional ByVal Append As Boolean = False) As Boolean
        Dim SW As System.IO.StreamWriter
        Dim fname As String
        Try
            fname = HttpContext.Current.Server.MapPath(Filename)
            SW = New System.IO.StreamWriter(fname, Append)
            SW.WriteLine(TextData)
            SW.Close()
            Return True
        Catch ex As Exception
            Return False
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function WriteBinaryFile(ByVal Filename As String, ByVal BinaryData As Object, Optional ByVal FileMode As System.IO.FileMode = IO.FileMode.Create) As Boolean
        Dim BW As System.IO.BinaryWriter
        Dim fname As String
        Try
            fname = HttpContext.Current.Server.MapPath(Filename)
            BW = New System.IO.BinaryWriter(System.IO.File.Open(fname, IO.FileMode.Create))
            BW.Write(BinaryData)
            BW.Close()
            Return True
        Catch ex As Exception
            Return False
        Finally
            BW = Nothing
        End Try
    End Function

    Public Sub SetCascadingListValue(ByRef C As AjaxControlToolkit.CascadingDropDown, ByVal SelectedValue As Object)
        Try
            C.SelectedValue = SelectedValue
        Catch ex As Exception
        End Try
    End Sub

    Public Function RemoveDuplicate(ByVal items As String(), Optional ByVal sort As Boolean = True) As String()
        Dim noDups As New ArrayList()
        For i As Integer = 0 To items.Length - 1
            If Not noDups.Contains(items(i).Trim()) Then
                noDups.Add(items(i).Trim())
            End If
        Next
        If sort Then
            noDups.Sort()
        End If
        'sorts list alphabetically 
        Dim uniqueItems As String() = New String(noDups.Count - 1) {}
        noDups.CopyTo(uniqueItems)
        Return uniqueItems
    End Function

    Public Sub BindDataList(ByRef DL As DataList, ByRef dataSource As Object, Optional ByVal AddBlankRow As Boolean = False)
        Dim DS As DataSet
        Dim DT As DataTable
        Dim DV As DataView

        Try
            If Not IsNothing(DL) AndAlso Not IsNothing(dataSource) Then
                Select Case TypeName(dataSource).ToUpper()
                    Case "DATASET"
                        DS = CType(dataSource, DataSet)
                        If DS.Tables.Count > 0 Then
                            DT = DS.Tables(0)
                            If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                                DT.Rows.Add(DT.NewRow)     ' Add blank row
                            End If
                            DL.DataSource = DT
                            DL.DataBind()
                        End If
                    Case "DATATABLE"
                        DT = CType(dataSource, DataTable)
                        If AddBlankRow AndAlso DT.Rows.Count = 0 Then
                            DT.Rows.Add(DT.NewRow)     ' Add blank row
                        End If
                        DL.DataSource = DT
                        DL.DataBind()
                    Case "DATAVIEW"
                        DV = CType(dataSource, DataView)
                        If AddBlankRow AndAlso DV.Count = 0 Then
                            DV.AddNew()     ' Add blank row
                        End If
                        DL.DataSource = DV
                        DL.DataBind()
                End Select
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub LoadMonthAbbrCombo(ByRef C As DropDownList, Optional ByVal IncBlank As Boolean = False, Optional ByVal DefaultMonth As String = "", Optional ByVal LangType As String = "EN", Optional ByVal BlankText As String = "")
        Dim I As Integer

        C.Items.Clear()
        If IncBlank Then
            C.Items.Add(New ListItem(BlankText, ""))
        End If
        If LangType & "" = "TH" Then
            For I = 1 To 12
                C.Items.Add(New ListItem(TMonth2(I), CStr(I).PadLeft(2, "0")))
            Next
        Else
            For I = 1 To 12
                C.Items.Add(New ListItem(EMonth2(I), CStr(I).PadLeft(2, "0")))
            Next
        End If

        If DefaultMonth <> "" Then
            C.SelectedValue = DefaultMonth
        Else
            C.SelectedIndex = -1
        End If
    End Sub

    Public Function ReadFileData(ByVal Filename As String) As Byte()
        Try
            'Dim fs As New System.IO.FileStream(HttpContext.Current.Server.MapPath(Filename), IO.FileMode.Open, IO.FileAccess.Read)
            Dim fs As New System.IO.FileStream(Filename, IO.FileMode.OpenOrCreate, IO.FileAccess.Read)
            Dim Buffer() As Byte = Nothing

            If Not IsNothing(fs) Then
                ReDim Buffer(fs.Length)
                fs.Read(Buffer, 0, fs.Length)
                fs.Close()
                ClearObject(fs)
            End If

            Return Buffer
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Created By Aoy 22/06/2552
    Public Function GetDBFieldType(ByVal Type As String) As DBUTIL.FieldTypes
        Select Case Type
            Case "TEXT" : Return DBUTIL.FieldTypes.ftText
            Case "NUMERIC" : Return DBUTIL.FieldTypes.ftNumeric
            Case "DATE" : Return DBUTIL.FieldTypes.ftDate
            Case "DATETIME" : Return DBUTIL.FieldTypes.ftDateTime
            Case "BINARY" : Return DBUTIL.FieldTypes.ftBinary
            Case Else
                Return DBUTIL.FieldTypes.ftText
        End Select
    End Function

    Public Sub TimeSpanToDate(ByVal d1 As DateTime, ByVal d2 As DateTime, ByRef years As Integer, ByRef months As Integer, ByRef days As Integer)
        ' compute & return the difference of two dates, 
        ' returning years, months & days 
        ' d1 should be the larger (newest) of the two dates 
        ' we want d1 to be the larger (newest) date 
        ' flip if we need to 
        If d1 < d2 Then
            Dim d3 As DateTime = d2
            d2 = d1
            d1 = d3
        End If

        ' compute difference in total months 
        months = 12 * (d1.Year - d2.Year) + (d1.Month - d2.Month)

        ' based upon the 'days', 
        ' adjust months & compute actual days difference 
        If d1.Day < d2.Day Then
            months -= 1
            days = DateTime.DaysInMonth(d2.Year, d2.Month) - d2.Day + d1.Day
        Else
            days = d1.Day - d2.Day
        End If
        ' compute years & actual months 
        years = months / 12
        months -= years * 12
    End Sub

    Public Function HaveSpecialChar(ByVal str As String) As Boolean
        Dim ChkSpecialChar As Boolean = False
        If str <> "" Then
            If str.IndexOf(">") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("<") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("..") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("--") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("`") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("'") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("|") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("&") <> -1 Then ChkSpecialChar = True
            If str.IndexOf(":") <> -1 Then ChkSpecialChar = True
            If str.IndexOf(";") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("$") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("@") <> -1 Then ChkSpecialChar = True
            If str.IndexOf(",") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("\'") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("\""") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("+") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("(") <> -1 Then ChkSpecialChar = True
            If str.IndexOf(")") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("""") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("=") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("<>") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("()") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("#") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("%") <> -1 Then ChkSpecialChar = True
            If str.IndexOf("*") <> -1 Then ChkSpecialChar = True
        End If
        Return ChkSpecialChar
    End Function
End Module

