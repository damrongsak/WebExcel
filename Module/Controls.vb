#Region ".NET Framework Class Import"
Imports System.Security
Imports System.Security.Principal
Imports System.Threading.Thread
Imports System.Net.Mail
Imports System.Data
#End Region


Public Module Controls

    Public Sub SetFileDisplayFormat(ByRef obj As Object, ByVal Path As String, ByVal FileName As String, Optional ByVal isFrontEnd As Boolean = False)
        Dim FullFilePath As String
        Dim FileType As String = ""
        Dim txt As String = ""
        Dim rndno As String

        Try
            If FileName <> "" Then
                rndno = Rnd()

                FileType = GetFileType(FileName)
                FullFilePath = gFilePath & Path & FileName & "?" & rndno
                Select Case True
                    Case InStr("|" & imgFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><img src='" & FullFilePath & "' border='0'>"
                        txt = "<img src='" & FullFilePath & "' border='0'>"
                    Case InStr("|" & clipFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object id='Player' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='1'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                        txt = "<object id='Player' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='0'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                    Case InStr("|" & soundFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object id='Player' width='220' height='45' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='1'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                        txt = "<object id='Player' width='220' height='45' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='0'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                    Case InStr("|" & flashFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0'><param name='movie' value='" & FullFilePath & "'><param name='quality' value='high'><embed src='" & FullFilePath & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash'></embed></object>"
                        txt = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0'><param name='movie' value='" & FullFilePath & "'><param name='quality' value='high'><embed src='" & FullFilePath & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash'></embed></object>"
                    Case Else 'docFileType
                        'txt = "<br><a href='" & FullFilePath & "' target='_blank'><b>Uploaded File</b></a>"
                        txt = "<a href='" & FullFilePath & "' target='_blank'><b>Uploaded File</b></a>"
                End Select
            End If

            Select Case obj.GetType().Name
                Case "Label" : CType(obj, Label).Text = txt
                Case Else : obj = txt
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Aoy 24/03/2552
    'Created 24/03/2552 Updated 24/03/2552
    Public Sub SetFileDisplayFormat2(ByRef obj As Object, ByVal Path As String, ByVal FileName As String, Optional ByVal LinkName As String = "", Optional ByVal ShowFileType As Boolean = True, Optional ByVal isFrontEnd As Boolean = False)
        Dim FullFilePath As String
        Dim FileType As String = ""
        Dim txt As String = ""
        Dim rndno As String

        Try
            If FileName <> "" Then
                rndno = Rnd()
                FileType = IIf(ShowFileType, GetFileType(FileName), "")
                FullFilePath = gFilePath & Path & FileName & "?" & rndno
                Select Case True
                    Case InStr("|" & imgFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><img src='" & FullFilePath & "' border='0'>"
                        txt = "<img src='" & FullFilePath & "' border='0'>"
                    Case InStr("|" & clipFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object id='Player' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='1'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                        txt = "<object id='Player' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='0'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                    Case InStr("|" & soundFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object id='Player' width='220' height='45' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='1'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                        txt = "<object id='Player' width='220' height='45' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT><PARAM NAME='URL' VALUE='" & FullFilePath & "'><PARAM NAME='playCount' VALUE='1'><PARAM NAME='autoStart' VALUE='0'><PARAM NAME='volume' VALUE='50'><PARAM NAME='uiMode' VALUE='full'></object>"
                    Case InStr("|" & flashFileType & "|", "|" & FileType & "|") > 0
                        'txt = "<br><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0'><param name='movie' value='" & FullFilePath & "'><param name='quality' value='high'><embed src='" & FullFilePath & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash'></embed></object>"
                        txt = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0'><param name='movie' value='" & FullFilePath & "'><param name='quality' value='high'><embed src='" & FullFilePath & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash'></embed></object>"
                    Case Else 'docFileType
                        'txt = "<br><a href='" & FullFilePath & "' target='_blank'><b>Uploaded File</b></a>"
                        txt = "<a href='" & FullFilePath & "' target='_blank'><b>" & IIf(LinkName <> "", LinkName, "Uploaded File") & "</b></a>"
                End Select
            End If

            Select Case obj.GetType().Name
                Case "Label" : CType(obj, Label).Text = txt
                Case Else : obj = txt
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Aoy"
    Public Function LoadCheckboxList(ByVal dtData As DataTable, ByVal ctrlName As String, ByVal fText As String, ByVal fValue As String, ByVal RepeatClNum1 As Integer, ByVal RepeatClNum2 As Integer, Optional ByVal BLevel As Boolean = False) As String
        Dim strChkList As String = ""
        Dim i1, i2, r1 As Integer
        Dim Row As DataRow
        Try
            i1 = 1 : i2 = 1 : r1 = 0
            If Not dtData Is Nothing Then
                For Each Row In dtData.Rows
                    If Row("LEVEL") & "" = "1" Then
                        'If i1 = 1 Then
                        '    strChkList &= "<table border=""0"" cellspacing=""2"" cellpadding=""0""><tr>"
                        'End If
                        'strChkList &= "<td align=""left""><input type=""checkbox"" name=""" & ctrlName & Row(fValue) & """ value=""" & Row(fValue) & """/>" & IIf(BLevel, "<b>", "") & Row(fText) & IIf(BLevel, "</b>", "") & "</td>"
                        'If i1 Mod RepeatClNum2 = 0 Then
                        '    strChkList &= "</tr>"
                        '    'If Not ((r1 + 1) = dtData.Rows.Count) Then
                        '    '    strChkList &= "<tr>"
                        '    'End If
                        'End If
                        'If getMaxID(dtData, Row("LEVEL") & "", fText, fValue) = Row(fValue) & "" And (r1 + 1) = dtData.Rows.Count Then
                        '    strChkList &= "</table>"
                        '    'Else
                        '    '    strChkList &= "<tr>"
                        'End If
                        'i1 = i1 + 1


                        If i1 = 1 Then
                            strChkList &= "<table border=""0"" cellspacing=""2"" cellpadding=""0""><tr>"
                        End If
                        strChkList &= "<td align=""left""><input type=""checkbox"" name=""" & ctrlName & Row(fValue) & """ value=""" & Row(fValue) & """/>" & IIf(BLevel, "<b>", "") & Row(fText) & IIf(BLevel, "</b>", "") & "</td>"
                        If i1 Mod RepeatClNum1 = 0 Then
                            If getMaxID(dtData, Row("LEVEL") & "", fText, fValue) = Row(fValue) & "" Then
                                If (r1 + 1) = dtData.Rows.Count Then
                                    strChkList &= "</tr></table>"
                                Else
                                    strChkList &= "</tr><tr>"
                                End If
                            Else
                                strChkList &= "</tr><tr>"
                            End If
                        End If
                        i1 = i1 + 1
                    Else
                        If i2 = 1 Then
                            strChkList &= "<td><table border=""0"" cellspacing=""2"" cellpadding=""0""><tr>"
                        End If
                        strChkList &= "<td align=""left""><input type=""checkbox"" name=""" & ctrlName & Row(fValue) & """ value=""" & Row(fValue) & """/>" & Row(fText) & "</td>"
                        If i2 Mod RepeatClNum2 = 0 Then
                            strChkList &= "</tr>"
                            If Not ((r1 + 1) = dtData.Rows.Count) Then
                                strChkList &= "<tr>"
                            End If
                        End If
                        If getMaxID(dtData, Row("LEVEL") & "", fText, fValue) = Row(fValue) & "" And (r1 + 1) = dtData.Rows.Count Then
                            If (r1 + 1) = dtData.Rows.Count Then
                                strChkList &= "</table></td></tr></table>"
                            Else
                                strChkList &= "</table></td></tr>"
                            End If
                        End If
                        i2 = i2 + 1
                    End If
                    r1 = r1 + 1
                Next
            End If
            Return strChkList
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Created By Aoy 12/05/2552 Updated By
    Public Function GetCheckboxListValue(ByRef ChkList As CheckBoxList) As String
        Dim LItem As ListItem = Nothing
        Dim CheckedValue As String = ""
        Try
            For Each LItem In ChkList.Items
                If LItem.Value <> "" And LItem.Selected Then
                    CheckedValue &= LItem.Value & ","
                End If
            Next
            Return CheckedValue
        Catch ex As Exception
            Return CheckedValue
        End Try
    End Function

    Private Function getMaxID(ByVal dtData As DataTable, ByVal Level As String, ByVal fText As String, ByVal fValue As String) As String
        Dim strMaxID As String = ""
        Dim DR As DataRow()
        Try
            DR = dtData.Select("LEVEL=" & Level, fValue & " DESC")
            If Not DR Is Nothing Then
                strMaxID = DR(0)(fValue).ToString()
            End If
            Return strMaxID
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub LoadCheckboxList(ByRef ChkList As CheckBoxList, ByVal DT As DataTable, ByVal fText As String, ByVal fValue As String, Optional ByVal BLevel As Boolean = False, Optional ByVal ClmRepeat1 As Integer = 1, Optional ByVal ClmRepeat2 As Integer = 3)
        Dim Row As DataRow
        Dim i As Integer
        Try
            ChkList.RepeatColumns = ClmRepeat2
            If Not DT Is Nothing Then
                For Each Row In DT.Rows
                    If Row("LEVEL") & "" = "1" Then
                        ChkList.Items.Add(New ListItem(IIf(BLevel, "<b>", "") & Row(fText) & IIf(BLevel, "</b>", ""), Row(fValue) & ""))
                        If ClmRepeat1 < ChkList.RepeatColumns Then
                            For i = 0 To (ChkList.RepeatColumns)
                                ChkList.Items.Add(New ListItem())
                                i = i + 1
                            Next
                        End If
                    Else
                        ChkList.Items.Add(New ListItem(Row(fText) & "", Row(fValue) & ""))
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub SetCheckedData(ByRef ChkList As CheckBoxList, ByVal ChkName As String, Optional ByVal strChecked As String = "")
        Dim listItem As ListItem = Nothing
        'Dim chk As HtmlInputCheckBox = Nothing
        Dim i As Integer
        Try
            i = 0
            For Each listItem In ChkList.Items
                If listItem.Value = "" Then
                    ChkList.Items(i).Attributes.Add("style", "visibility:hidden;")
                Else
                    listItem.Selected = IsCheckedData(listItem.Value, strChecked)
                    'listItem.Selected = True
                End If
                i = i + 1
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function IsCheckedData(ByVal Value As String, ByVal strChecked As String) As Boolean
        Dim Checked As String = ""
        Try
            If strChecked <> "" Then
                Checked = "," & strChecked
                If Checked.IndexOf("," & Value & ",") = -1 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GenCtrl(ByVal ColumnType As String, ByVal CtrlType As String, ByVal CtrlID As String _
    , ByVal DataType As String, Optional ByVal Value As Object = Nothing _
    , Optional ByVal Width As Object = Nothing, Optional ByVal IsAddCtrl As Boolean = False _
    , Optional ByVal IsLabel As Boolean = True, Optional ByVal IsSearch As Boolean = False _
    , Optional ByVal Value2 As Object = Nothing, Optional ByVal EditFlag As String = "Y") As String
        Dim Ctrl As String = ""
        Dim DT As DataTable = Nothing
        Dim DR As DataRow
        Dim CtrlName As String = ""
        Try
            DataType = Trim(DataType)
            Select Case DataType
                Case "DATETIME"
                    If TypeName(Value).ToLower <> "string" Then
                        If CtrlType = "TEXTBOX" Then
                            Value = FormatDate(Value, "DD/MM/YYYY") & ""
                        Else
                            Value = FormatDate(Value, "DD/MM/YYYY HH:MIN") & ""
                        End If
                    End If
                Case "DATE"
                    If TypeName(Value).ToLower <> "string" Then
                        Value = FormatDate(Value, "DD/MM/YYYY") & ""
                    End If
                Case Else
                    Value = Value & ""
            End Select
            If IsAddCtrl Then
                CtrlName = "Add"
            ElseIf IsSearch Then
                CtrlName = "Search"
            End If
            CtrlType = Trim(CtrlType)
            If EditFlag = "Y" OrElse IsSearch Then
                Select Case CtrlType
                    Case "TEXTBOX", "YEAR", "NUMBER" : Ctrl = "<input id=""txt" & CtrlName & CtrlID & """ name=""txt" & CtrlName & "" & CtrlID & """ type=""text"" value=""" & Value & """ " & IIf(Width <> "", "style=""width:" & Width & "px""", "") & " />"
                    Case "TEXTAREA"
                        If IsSearch Then
                            Ctrl = "<input id=""txt" & CtrlName & CtrlID & """ name=""txt" & CtrlName & "" & CtrlID & """ type=""text"" value=""" & Value & """ " & IIf(Width <> "", "style=""width:" & Width & "px""", "") & " />"
                        Else
                            Ctrl = "<textarea id=""txt" & CtrlName & CtrlID & """ name=""txt" & IIf(IsAddCtrl, "Add", "") & "" & CtrlID & """  rows=""5"" " & IIf(Width <> "", "style=""width:" & Width & "px""", "") & ">" & Value & "</textarea>"
                        End If
                    Case "DROPDOWN"
                        DR = GetDR(DAL.SearchConfigColumnTypes(ColumnType))
                        If Not IsNothing(DR) Then
                            If DR("PARAM1") & "" <> "" AndAlso DR("PARAM2") & "" <> "" AndAlso DR("PARAM3") & "" <> "" Then
                                DT = DAL.QueryData("SELECT * FROM " & DR("PARAM1") & " ORDER BY " & DR("PARAM3") & "")
                                Ctrl = GenDropDownCtrl(DT, CtrlID, DR("PARAM2") & "", DR("PARAM3") & "", Value, IsAddCtrl, IsSearch, Width)
                            End If
                        End If
                    Case "DATE"
                        If IsSearch Then
                            'Ctrl = "<input id=""txtDateF" & CtrlName & CtrlID & """ name=""txtDateF" & CtrlName & CtrlID & """ type=""text"" value=""" & Value & """ onclick=""popUpCalendar(this,document.getElementById('txtDateF" & CtrlName & CtrlID & "'), 'dd/mm/yyyy', -120)"" style=""width:80px"" /> - "
                            'Ctrl &= "<input id=""txtDateT" & CtrlName & CtrlID & """ name=""txtDateT" & CtrlName & CtrlID & """ type=""text"" value=""" & Value2 & """ onclick=""popUpCalendar(this,document.getElementById('txtDateT" & CtrlName & CtrlID & "'), 'dd/mm/yyyy', -120)"" style=""width:80px"" />"
                            Ctrl = "<input id=""txtDateF" & CtrlName & CtrlID & """ name=""txtDateF" & CtrlName & CtrlID & """ type=""text"" value=""" & Value & """ style=""width:80px"" /> - "
                            Ctrl &= "<input id=""txtDateT" & CtrlName & CtrlID & """ name=""txtDateT" & CtrlName & CtrlID & """ type=""text"" value=""" & Value2 & """ style=""width:80px"" />"
                        Else
                            'Ctrl = "<input id=""txtDate" & CtrlName & CtrlID & """ name=""txtDate" & CtrlName & CtrlID & """ type=""text"" value=""" & Value & """ onclick=""popUpCalendar(this,document.getElementById('txtDate" & CtrlName & CtrlID & "'), 'dd/mm/yyyy', -120)"" style=""width:80px"" />"
                            Ctrl = "<input id=""txtDate" & CtrlName & CtrlID & """ name=""txtDate" & CtrlName & CtrlID & """ type=""text"" value=""" & Value & """ style=""width:80px""/>"
                        End If
                    Case Else
                        If IsLabel And Not IsSearch And Not IsAddCtrl And EditFlag <> "Y" Then
                            Ctrl = Value
                        Else
                            Ctrl = "<input id=""txt" & CtrlName & CtrlID & """ name=""txt" & CtrlName & CtrlID & """ type=""text"" value=""" & Value & """ " & IIf(Width <> "", "style=""width:" & Width & "px""", "") & " />"
                        End If
                End Select
            Else
                Ctrl = Value
            End If
            Return Ctrl
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function GetCtrlValue(ByVal ColumnType As String, ByVal CtrlType As String, ByVal CtrlID As String _
    , ByVal DataType As String, Optional ByVal IsAddCtrl As Boolean = False) As String
        Dim Value As String = ""
        Dim CtrlName As String = ""
        Try
            CtrlType = Trim(CtrlType)
            Select Case CtrlType
                Case "TEXTBOX", "TEXTAREA", "YEAR", "NUMBER" : CtrlName = "txt" & IIf(IsAddCtrl, "Add", "") & "" & CtrlID
                Case "DATE" : CtrlName = "txtDate" & IIf(IsAddCtrl, "Add", "") & "" & CtrlID
                Case "DROPDOWN" : CtrlName = "ddl" & IIf(IsAddCtrl, "Add", "") & "" & CtrlID
                Case Else
                    CtrlName = "txt" & IIf(IsAddCtrl, "Add", "") & "" & CtrlID
            End Select
            If CtrlName <> "" Then
                Value = ValidateData(HttpContext.Current.Request.Form(CtrlName) & "")
                If CtrlType = "YEAR" Then
                    If CInt(Value) < 2500 Then
                        Value = CStr(CInt(Value) + 543)
                    End If
                End If
            End If
        Catch ex As Exception
            Value = ""
        End Try
        Return Value
    End Function

    Public Sub GetCtrlValue2(ByRef Value1 As String, ByRef Value2 As String, ByVal ColumnType As String, ByVal CtrlType As String, ByVal CtrlID As String _
    , ByVal DataType As String, Optional ByVal IsAddCtrl As Boolean = False, Optional ByVal IsSearch As Boolean = False)
        Dim CtrlName As String = "", Ctrl1 As String = "", Ctrl2 As String = ""
        Try
            Value1 = "" : Value2 = ""
            CtrlType = Trim(CtrlType)

            If IsAddCtrl Then
                CtrlName = "Add"
            ElseIf IsSearch Then
                CtrlName = "Search"
            End If
            Select Case CtrlType
                Case "TEXTBOX", "TEXTAREA", "YEAR", "NUMBER" : Ctrl1 = "txt" & CtrlName & CtrlID
                Case "DATE"
                    If IsSearch Then
                        Ctrl1 = "txtDateF" & CtrlName & CtrlID
                        Ctrl2 = "txtDateT" & CtrlName & CtrlID
                    Else
                        Ctrl1 = "txtDate" & CtrlName & CtrlID
                    End If
                Case "DROPDOWN" : Ctrl1 = "ddl" & CtrlName & CtrlID
                Case Else
                    If IsSearch Then
                        Ctrl1 = "txt" & CtrlName & CtrlID
                    End If
            End Select
            If Ctrl1 <> "" Then
                Value1 = ValidateData(HttpContext.Current.Request.Form(Ctrl1) & "")
                If CtrlType = "YEAR" Then
                    If CInt(Value1) < 2500 Then
                        Value1 = CStr(CInt(Value1) + 543)
                    End If
                End If
            End If
            If Ctrl2 <> "" Then
                Value2 = ValidateData(HttpContext.Current.Request.Form(Ctrl2) & "")
                If CtrlType = "YEAR" Then
                    If CInt(Value2) < 2500 Then
                        Value2 = CStr(CInt(Value2) + 543)
                    End If
                End If
            End If
        Catch ex As Exception
            Value1 = "" : Value2 = ""
        End Try
    End Sub

    Public Function GetValue(ByVal DataType As String, ByVal Value As Object) As String
        Dim retValue As String = ""
        Try
            DataType = Trim(DataType)
            Select Case DataType.ToUpper
                Case "DATE" : retValue = FormatDate(Value, "DD/MM/»»»»") & ""

                Case "DATETIME" : retValue = FormatDate(Value, "DD/MM/»»»» HH:MIN") & ""
                Case Else
                    retValue = Value & ""
            End Select
        Catch ex As Exception
            retValue = ""
        End Try
        Return retValue
    End Function

    Public Function GenDropDownCtrl(ByVal DT As DataTable, ByVal CtrlID As String _
    , ByVal FieldValue As String, ByVal FieldDesc As String, Optional ByVal ValSelected As String = "" _
    , Optional ByVal IsAddCtrl As Boolean = False, Optional ByVal IsSearch As Boolean = False _
    , Optional ByVal Width As String = "") As String
        Dim ddlCtrl As String = "", CtrlName As String = ""
        If IsAddCtrl Then
            CtrlName = "Add"
        ElseIf IsSearch Then
            CtrlName = "Search"
        End If
        ddlCtrl = "<select id=""ddl" & CtrlName & CtrlID & """ name=""ddl" & CtrlName & CtrlID & """" & IIf(Width <> "", " style=""width:" & Width & "px""", "") & ">"
        ddlCtrl &= "<option value="""" " & IIf(ValSelected = "", "selected=""selected""", "") & "></option>"
        For Each DR In DT.Rows
            ddlCtrl &= "<option value=""" & DR(FieldValue) & """ " & IIf(ValSelected = DR(FieldValue) & "", "selected=""selected""", "") & ">" & DR(FieldDesc) & "</option>"
        Next
        ddlCtrl &= "</select>"
        Return ddlCtrl
    End Function

#End Region

#Region "JSON"
    Public Function GenDropDownCtrlJson(ByVal DT As DataTable, ByVal CtrlID As String _
, ByVal FieldValue As String, ByVal FieldDesc As String, Optional ByVal ValSelected As String = "" _
, Optional ByVal IsAddCtrl As Boolean = False, Optional ByVal IsSearch As Boolean = False _
, Optional ByVal Width As String = "") As String
        Dim ddlCtrl As String = "", CtrlName As String = ""
        'If IsAddCtrl Then
        '    CtrlName = "Add"
        'ElseIf IsSearch Then
        '    CtrlName = "Search"
        'End If
        'ddlCtrl = "<select id=""ddl" & CtrlName & CtrlID & """ name=""ddl" & CtrlName & CtrlID & """" & IIf(Width <> "", " style=""width:" & Width & "px""", "") & ">"
        'ddlCtrl &= "<option value="""" " & IIf(ValSelected = "", "selected=""selected""", "") & "></option>"
        'For Each DR In DT.Rows
        '    ddlCtrl &= "<option value=""" & DR(FieldValue) & """ " & IIf(ValSelected = DR(FieldValue) & "", "selected=""selected""", "") & ">" & DR(FieldDesc) & "</option>"
        'Next
        'ddlCtrl &= "</select>"
        '{"Table" : 
        '[
        '{"stateid" : "2","statename" : "Tamilnadu"},
        '{"stateid" : "3","statename" : "Karnataka"},
        '{"stateid" : "4","statename" : "Andaman and Nicobar"},
        '{"stateid" : "5","statename" : "Andhra Pradesh"},
        '{"stateid" : "6","statename" : "Arunachal Pradesh"}
        ']
        '}
        ddlCtrl = "{'Table' :" & vbCrLf
        ddlCtrl &= "[" & vbCrLf
        For Each DR In DT.Rows
            ddlCtrl &= String.Format("{'{0}' : '{1}','{2}' : '{3}'}," & vbCrLf, FieldValue, DR(FieldValue), FieldDesc, DR(FieldDesc))
        Next
        ddlCtrl &= "]" & vbCrLf
        ddlCtrl &= "}" & vbCrLf
        Return ddlCtrl
    End Function
#End Region
End Module



