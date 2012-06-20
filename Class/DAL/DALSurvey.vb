Partial Public Class DALComponent
    Public Function SearchForm(Optional ByVal FormID As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "F.FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT F.FORM_ID , F.FORM_NAME, F.FORM_DESC,F.ACTIVE_FLAG,F.PUBLIC_FLAG,F.FORM_IMAGE, " & _
                 "COUNT(Q.FORM_ID) AS CNT " & _
                 "FROM FORMS F, QUESTIONS Q " & _
                 "WHERE F.FORM_ID = Q.FORM_ID(+) "

            If CriteriaSQL <> "" Then
                SQL += " AND " + CriteriaSQL
            End If
            SQL &= " GROUP BY F.FORM_ID,F.FORM_NAME, F.FORM_DESC,F.ACTIVE_FLAG,F.PUBLIC_FLAG,F.FORM_IMAGE "
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function SearchSection(Optional ByVal SectionID As String = "", Optional ByVal FormID As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
            If SectionID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT S.FORM_ID, S.SECTION_ID, S.SECTION_NO, S.SECTION_DESC, COUNT(Q.QUESTION_ID) AS CNT,F.FORM_NAME " + _
              "FROM SECTIONS S,QUESTIONS Q,FORMS F WHERE S.FORM_ID=Q.FORM_ID(+) AND S.SECTION_ID=Q.SECTION_ID(+) AND S.FORM_ID=F.FORM_ID(+) "


            If CriteriaSQL <> "" Then
                SQL += " AND " + CriteriaSQL
            End If
            SQL &= " GROUP BY S.FORM_ID, S.SECTION_ID, S.SECTION_NO, S.SECTION_DESC,F.FORM_NAME ORDER BY SECTION_NO"
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function SearchQuestion(Optional ByVal QuestionID As String = "", Optional ByVal SectionID As String = "", Optional ByVal FormID As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "Q.FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
            If SectionID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "Q.SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
            If QuestionID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "Q.QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT Q.*, D.QUESTION_TYPE_DESC " + _
                      ",S.SECTION_NO,S.SECTION_DESC,F.FORM_NAME FROM QUESTIONS Q, VW_QUESTION_TYPES D,SECTIONS S,FORMS F " + _
                      " WHERE Q.QUESTION_TYPE=D.QUESTION_TYPE(+) AND Q.SECTION_ID=S.SECTION_ID(+) AND Q.FORM_ID=F.FORM_ID(+)"
            If CriteriaSQL <> "" Then
                SQL += " AND " + CriteriaSQL
            End If
            SQL &= " ORDER BY TO_BINARY_DOUBLE(Q.QUESTION_NO)"
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    'Public Function SearchSurvey(Optional ByVal SurveyID As String = "", Optional ByVal FormID As String = "" _
    ', Optional ByVal BatchID As String = "", Optional ByVal CompanyName As String = "" _
    ', Optional ByVal Name As String = "", Optional ByVal SurName As String = "", Optional ByVal Age As String = "" _
    ', Optional ByVal Sex As String = "", Optional ByVal Education As String = "", Optional ByVal WorkType As String = "" _
    ', Optional ByVal FromDate As String = "", Optional ByVal ToDate As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
    '    Dim SQL As String
    '    Dim CriteriaSQL As String
    '    Dim DT As DataTable = Nothing

    '    Try
    '        CriteriaSQL = OtherCriteria
    '        If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
    '        If SurveyID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
    '        If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
    '        If CompanyName & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.COMPANY_NAME", CompanyName, DBUTIL.FieldTypes.ftText)
    '        If Name & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.NAME", Name, DBUTIL.FieldTypes.ftText)
    '        If SurName & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.SURNAME", SurName, DBUTIL.FieldTypes.ftText)
    '        If Age & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.AGE", Age, DBUTIL.FieldTypes.ftNumeric)
    '        If Sex & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.SEX", Sex, DBUTIL.FieldTypes.ftText)
    '        If Education & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.EDUCATION", Education, DBUTIL.FieldTypes.ftNumeric)
    '        If WorkType & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.WORK_TYPE", WorkType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddCriteriaRange(CriteriaSQL, "S.DATE_UPDATED", AppDateValue(FromDate), AppDateValue(ToDate), DBUTIL.FieldTypes.ftDate)

    '        SQL = "SELECT * " + _
    '          " FROM SURVEYS S,REF_AGE_DATA AD,REF_WORK_TYPES WT,REF_EDUCATION_DATA ED " + _
    '          " WHERE S.AGE=AD.AGE_ID(+) AND S.WORK_TYPE=WT.WORK_TYPE(+) AND S.EDUCATION=ED.EDUCATION_ID(+)"

    '        If CriteriaSQL <> "" Then
    '            SQL += " AND " + CriteriaSQL
    '        End If
    '        DB.OpenDT(DT, SQL)
    '        Return DT

    '    Catch ex As Exception
    '        Throw (ex)
    '    End Try
    'End Function

    Public Function SearchSurvey(Optional ByVal SurveyID As String = "", Optional ByVal FormID As String = "" _
   , Optional ByVal BatchID As String = "", Optional ByVal CompanyName As String = "" _
   , Optional ByVal Name As String = "", Optional ByVal SurName As String = "" _
   , Optional ByVal FromDate As String = "", Optional ByVal ToDate As String = "", Optional ByVal PositionName As String = "" _
   , Optional ByVal EMail As String = "", Optional ByVal TelNo As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
            If SurveyID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
            If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            If CompanyName & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(COMPANY_NAME)", CompanyName.ToUpper, DBUTIL.FieldTypes.ftText)
            If Name & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(NAME)", Name.ToUpper, DBUTIL.FieldTypes.ftText)
            If SurName & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(SURNAME)", SurName.ToUpper, DBUTIL.FieldTypes.ftText)
            If PositionName & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(POSITION_NAME)", PositionName.ToUpper, DBUTIL.FieldTypes.ftText)
            If EMail & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(EMAIL)", EMail.ToUpper, DBUTIL.FieldTypes.ftText)
            If TelNo & "" <> "" Then DB.AddCriteria(CriteriaSQL, "UPPER(TEL_NO)", TelNo.ToUpper, DBUTIL.FieldTypes.ftText)
            DB.AddCriteriaRange(CriteriaSQL, "DATE_UPDATED", AppDateValue(FromDate), AppDateValue(ToDate), DBUTIL.FieldTypes.ftDate)

            SQL = "SELECT * " + _
              " FROM SURVEYS "

            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function ManageForms(ByVal op As Integer, ByVal FormID As String _
    , Optional ByVal FormName As String = Nothing, Optional ByVal FormDesc As String = Nothing _
    , Optional ByVal FormImage As String = Nothing, Optional ByVal TotalScore As String = Nothing _
    , Optional ByVal ActiveFlag As String = Nothing, Optional ByVal PublicFlag As String = Nothing _
    , Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(FormID) Then DB.AddSQL(op, SQL1, SQL2, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If

                If Not IsNothing(FormName) Then DB.AddSQL(op, SQL1, SQL2, "FORM_NAME", FormName, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(FormDesc) Then DB.AddSQL(op, SQL1, SQL2, "FORM_DESC", FormDesc, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(FormImage) Then DB.AddSQL(op, SQL1, SQL2, "FORM_IMAGE", FormImage, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(TotalScore) Then DB.AddSQL(op, SQL1, SQL2, "TOTAL_SCORE", TotalScore, DBUTIL.FieldTypes.ftNumeric)
                If Not IsNothing(ActiveFlag) Then DB.AddSQL(op, SQL1, SQL2, "ACTIVE_FLAG", ActiveFlag, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(PublicFlag) Then DB.AddSQL(op, SQL1, SQL2, "PUBLIC_FLAG", PublicFlag, DBUTIL.FieldTypes.ftText)

            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "FORMS", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageQuestions(ByVal op As Integer, ByVal QuestionID As String _
    , Optional ByVal FormID As String = Nothing, Optional ByVal SectionID As String = Nothing _
    , Optional ByVal QuestionNo As String = Nothing, Optional ByVal QuestionDesc As String = Nothing _
    , Optional ByVal QuestionType As String = Nothing, Optional ByVal QuestionScore As String = Nothing _
    , Optional ByVal QuestionImage As String = Nothing, Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(QuestionID) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(FormID) Then DB.AddSQL(op, SQL1, SQL2, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(SectionID) Then DB.AddSQL(op, SQL1, SQL2, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If
                If Not IsNothing(QuestionNo) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_NO", QuestionNo, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(QuestionDesc) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_DESC", QuestionDesc, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(QuestionType) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_TYPE", QuestionType, DBUTIL.FieldTypes.ftNumeric)
                If Not IsNothing(QuestionScore) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_SCORE", QuestionScore, DBUTIL.FieldTypes.ftNumeric)
                If Not IsNothing(QuestionImage) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_IMAGE", QuestionImage, DBUTIL.FieldTypes.ftText)
                DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "QUESTIONS", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageSection(ByVal op As Integer, ByVal SectionID As String _
    , ByVal FormID As String, Optional ByVal SectionNo As String = Nothing _
    , Optional ByVal SectionDesc As String = Nothing, Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(FormID) Then DB.AddSQL(op, SQL1, SQL2, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(SectionID) Then DB.AddSQL(op, SQL1, SQL2, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If
                If Not IsNothing(SectionNo) Then DB.AddSQL(op, SQL1, SQL2, "SECTION_NO", SectionNo, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(SectionDesc) Then DB.AddSQL(op, SQL1, SQL2, "SECTION_DESC", SectionDesc, DBUTIL.FieldTypes.ftText)
                DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "SECTIONS", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ManageChoice(ByVal op As Integer, ByVal ChoiceID As String _
    , ByVal QuestionID As String, Optional ByVal ChoiceDesc As String = Nothing, Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "CHOICE_ID", ChoiceID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(QuestionID) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(ChoiceID) Then DB.AddSQL(op, SQL1, SQL2, "CHOICE_ID", ChoiceID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If
                If Not IsNothing(ChoiceDesc) Then DB.AddSQL(op, SQL1, SQL2, "CHOICE_DESC", ChoiceDesc, DBUTIL.FieldTypes.ftText)
            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "CHOICES", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchChoice(Optional ByVal ChoiceID As String = "", Optional ByVal QuestionID As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If ChoiceID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "CHOICE_ID", ChoiceID, DBUTIL.FieldTypes.ftNumeric)
            If QuestionID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM CHOICES "
            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            SQL &= " ORDER BY CHOICE_ID"
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function SearchQuestionType(Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria

            SQL = "SELECT * FROM VW_QUESTION_TYPES"
            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            SQL &= " ORDER BY QUESTION_TYPE"
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function ManageQuestion(ByVal op As Integer, ByVal QuestionID As String, ByVal SectionID As String _
    , ByVal FormID As String, Optional ByVal QuestionNo As String = Nothing, Optional ByVal QuesDesc As String = Nothing _
    , Optional ByVal QuesType As String = Nothing, Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
                DB.AddCriteria(Criteria, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(QuestionID) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_ID", QuestionID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(SectionID) Then DB.AddSQL(op, SQL1, SQL2, "SECTION_ID", SectionID, DBUTIL.FieldTypes.ftNumeric)
                    If Not IsNothing(FormID) Then DB.AddSQL(op, SQL1, SQL2, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If
                If Not IsNothing(QuestionNo) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_NO", QuestionNo, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(QuesDesc) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_DESC", QuesDesc, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(QuesType) Then DB.AddSQL(op, SQL1, SQL2, "QUESTION_TYPE", QuesType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "QUESTIONS", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Public Function SearchSurveyDetail(Optional ByVal SurveyID As String = "", Optional ByVal QuestionID As String = "" _
    ', Optional ByVal SectionID As String = "", Optional ByVal BatchID As String = "", Optional ByVal FormID As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
    '    Dim SQL, SQL2 As String
    '    Dim CriteriaSQL, CriteriaSQL2 As String
    '    Dim DT As DataTable = Nothing

    '    Try
    '        CriteriaSQL = OtherCriteria
    '        If SurveyID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.SURVEY_ID", SurveyID, DBUtil.FieldTypes.ftNumeric)
    '        If QuestionID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "S.QUESTION_ID", QuestionID, DBUtil.FieldTypes.ftNumeric)
    '        If SectionID & "" <> "" Then
    '            DB.AddCriteria(CriteriaSQL, "Q.SECTION_ID", SectionID, DBUtil.FieldTypes.ftNumeric)
    '            DB.AddCriteria(CriteriaSQL2, "SECTION_ID", SectionID, DBUtil.FieldTypes.ftNumeric)
    '        End If

    '        If BatchID & "" <> "" Then
    '            DB.AddCriteria(CriteriaSQL, "SD.BATCH_ID", BatchID, DBUtil.FieldTypes.ftNumeric)
    '        End If
    '        If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL2, "FORM_ID", FormID, DBUtil.FieldTypes.ftNumeric)

    '        SQL = "SELECT Q.QUESTION_ID, Q.FORM_ID, Q.SECTION_ID, Q.QUESTION_NO, Q.QUESTION_DESC, Q.QUESTION_TYPE" & _
    '        ", SD.SURVEY_ID, SD.CHOICE_ID, SD.ANSWER,SD.REMARK FROM QUESTIONS Q INNER JOIN " & _
    '        "SURVEY_DETAILS SD ON Q.QUESTION_ID = SD.QUESTION_ID "
    '        SQL2 = "SELECT Q.QUESTION_ID FROM QUESTIONS Q LEFT OUTER JOIN " & _
    '        "SURVEY_DETAILS SD ON Q.QUESTION_ID = SD.QUESTION_ID LEFT OUTER JOIN SURVEYS S "
    '        If CriteriaSQL <> "" Then
    '            SQL &= " WHERE " + CriteriaSQL
    '            SQL2 &= " WHERE " + CriteriaSQL
    '        End If
    '        If SQL2 <> "" Then
    '            SQL &= " UNION "
    '            SQL &= "SELECT QUESTION_ID, FORM_ID, SECTION_ID, QUESTION_NO, QUESTION_DESC, QUESTION_TYPE" & _
    '            ", NULL AS SURVEY_ID,  NULL AS CHOICE_ID,  NULL AS ANSWER, NULL AS REMARK FROM QUESTIONS WHERE QUESTION_ID NOT IN (" & SQL2 & ")"
    '            If CriteriaSQL2 <> "" Then SQL &= " AND " & CriteriaSQL2
    '        End If
    '        SQL = "SELECT * FROM (" & SQL & ") ORDER BY TO_BINARY_DOUBLE(QUESTION_NO)"
    '        HttpContext.Current.Response.Write(SQL)
    '        DB.OpenDT(DT, SQL)
    '        Return DT
    '    Catch ex As Exception

    '    End Try
    'End Function

    Public Function SearchBatch(Optional ByVal BatchID As String = "", Optional ByVal FormID As String = "" _
    , Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT B.*,F.FORM_NAME,F.FORM_IMAGE FROM BATCHES B,FORMS F WHERE B.FORM_ID=F.FORM_ID(+)"
            If CriteriaSQL <> "" Then
                SQL += " AND " + CriteriaSQL
            End If
            SQL &= " ORDER BY B.DATE_UPDATED DESC"
            'HttpContext.Current.Response.Write(SQL)
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    'Public Function SaveSurvey(ByRef SurveyID As String, ByVal BatchID As String _
    ', ByVal Name As String, ByVal SurName As String, ByVal Sex As String, ByVal Age As String _
    ', ByVal Education As String, ByVal WorkType As String, ByVal CompanyName As String _
    ', ByVal Address1 As String, ByVal Address2 As String, ByVal Address3 As String _
    ', ByVal Address4 As String, ByVal Tumbon As String, ByVal Amphur As String _
    ', ByVal Province As String, ByVal ZipCode As String, ByVal TelNo As String _
    ', ByVal EMail As String, ByVal WorkTypeOther As String, ByVal Request As HttpRequest _
    ', ByVal controlPassed5 As String(), ByVal controlPassed6 As String) As String
    '    Dim SQL1, SQL2, SQL As String
    '    Dim MaxID As Long
    '    Dim SysFileName As String
    '    Dim op As Integer
    '    Dim Obj As Object
    '    Dim keyItemID, VALUE, controlPassed, tmpSQL As String
    '    Dim QID, CID, tmpCID, i, r As Integer
    '    Dim statusQ5, statusT5 As Boolean
    '    'Dim controlPassed5(), controlPassed6 As String

    '    Try
    '        If SurveyID <> "0" Then
    '            DB.ExecSQL("DELETE FROM SURVEYS WHERE SURVEY_ID=" & SurveyID)
    '            DB.ExecSQL("DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID=" & SurveyID)
    '        Else
    '            SurveyID = ToInt("0" & LookupSQL("SELECT MAX(SURVEY_ID) FROM SURVEYS")) + 1
    '        End If

    '        op = opINSERT

    '        r = 0
    '        controlPassed = ""
    '        For Each Obj In Request.Form

    '            SQL = "" : SQL1 = "" : SQL2 = "" : VALUE = ""
    '            tmpSQL = "" : tmpCID = -1
    '            QID = -1 : CID = -1
    '            statusT5 = True : statusQ5 = True ' ให้สามารถเข้าไปเก็บชื่อใน controlPassed5() ได้

    '            Select Case Mid(CStr(Obj), 1, 4)
    '                Case "txt_"
    '                    If Request.Form(CStr(Obj)) & "" <> "" Then
    '                        QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
    '                        'CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                        If CStr(Obj).IndexOf("C") = CStr(Obj).Length - 1 Then
    '                            CID = 0
    '                        Else
    '                            CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                        End If
    '                        VALUE = Request.Form(CStr(Obj)) & "" 'Replace(Request.Form(CStr(Obj)) & "", vbCrLf, "<BR>")
    '                    End If
    '                Case "rdb_"
    '                    QID = CInt(Mid(CStr(Obj), 6))
    '                    CID = CInt(Mid(Request.Form(CStr(Obj)), Request.Form(CStr(Obj)).IndexOf("C") + 2))
    '                Case "chk_"
    '                    QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
    '                    CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                Case "rdt_" ' เป็น radio ที่มี textbox ด้วย 
    '                    Select Case Mid(CStr(Obj), 1, 5)
    '                        Case "rdt_T" ' ถ้ามี textbox แสดงว่าต้องมีการติ๊กด้วย
    '                            If Request.Form(CStr(Obj)) & "" <> "" Then ' user กรอกค่าใน txtbox
    '                                QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
    '                                CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                                'QID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("Q") + 2, CStr(Request.Form(CStr(Obj))).IndexOf("C") - CStr(Request.Form(CStr(Obj))).IndexOf("Q") - 1))
    '                                'CID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("C") + 2))
    '                                VALUE += Request.Form(CStr(Obj))
    '                                For i = 0 To controlPassed5.Length - 1
    '                                    If InStr(controlPassed5(i), "rdt_Q" & QID) > 0 Then ' ถ้ามีการ save chkbox ไปก่อนแล้ว
    '                                        tmpCID = CInt(Mid(controlPassed5(i), controlPassed5(i).IndexOf("C") + 2))
    '                                        tmpSQL = "DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID = " & SurveyID & " and QUESTION_ID = " & QID & " AND CHOICE_ID = " & tmpCID
    '                                        DB.ExecSQL(tmpSQL)
    '                                        controlPassed5(i) = "chk_Q" & QID & "C" & CID & ""
    '                                        statusT5 = False
    '                                    End If
    '                                Next
    '                                If statusT5 Then
    '                                    controlPassed5(r) = "chk_Q" & QID & "C" & CID & ""
    '                                    r += 1
    '                                End If
    '                            End If

    '                        Case "rdt_Q"
    '                            For i = 0 To controlPassed5.Length - 1
    '                                If InStr(controlPassed5(i), "Q" & CInt(Mid(CStr(Obj), 6))) <> 0 Then ' control ถูก Save ลงไน DB แล้ว
    '                                    statusQ5 = False ' เคย save ค่านี้แล้ว
    '                                End If
    '                            Next
    '                            If statusQ5 Then
    '                                QID = CInt(Mid(CStr(Obj), 6))
    '                                CID = CInt(Mid(Request.Form(CStr(Obj)), Request.Form(CStr(Obj)).IndexOf("C") + 2))

    '                                controlPassed5(r) = Request.Form(CStr(Obj))
    '                                r += 1
    '                            End If
    '                    End Select
    '                Case "cht_"
    '                    Select Case Mid(CStr(Obj), 1, 5)
    '                        Case "cht_T" ' ถ้ามี textbox แสดงว่าต้องมีการติ๊กด้วย
    '                            If Request.Form(CStr(Obj)) & "" <> "" Then ' user กรอกค่าใน txtbox
    '                                'QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
    '                                'CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                                QID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("Q") + 2, CStr(Request.Form(CStr(Obj))).IndexOf("C") - CStr(Request.Form(CStr(Obj))).IndexOf("Q") - 1))
    '                                CID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("C") + 2))
    '                                VALUE += Request.Form(CStr(Obj))
    '                                If InStr(controlPassed, "cht_Q" & QID & "C" & CID) > 0 Then ' ถ้ามีการ save chkbox ไปก่อนแล้ว
    '                                    tmpSQL = "DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID = " & SurveyID & " and QUESTION_ID = " & QID & " AND CHOICE_ID = " & CID
    '                                    DB.ExecSQL(tmpSQL)
    '                                    controlPassed += "," & "chk_TQ" & QID & "C" & CID
    '                                Else ' ยังไม่มีการ save control ที่ผ่านไปแล้ว
    '                                    controlPassed += "," & "chk_Q" & QID & "C" & CID & "," & "chk_TQ" & QID & "C" & CID
    '                                End If
    '                            End If
    '                        Case "cht_Q" ' ถ้ามีการติ๊ก ต้องเช็กว่ามี textbox ด้วยรึเปล่า
    '                            If InStr(controlPassed, CStr(Obj)) = 0 Then ' control ยังไม่ถูก Save ลงไน DB
    '                                QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
    '                                CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
    '                            End If
    '                            controlPassed += "," & CStr(Obj)
    '                    End Select

    '            End Select
    '            If QID <> -1 And CID <> -1 Then
    '                DB.AddSQL(op, SQL1, SQL2, "SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddSQL(op, SQL1, SQL2, "QUESTION_ID", QID, DBUTIL.FieldTypes.ftNumeric)
    '                DB.AddSQL(op, SQL1, SQL2, "CHOICE_ID", CID, DBUTIL.FieldTypes.ftNumeric)

    '                If VALUE <> "" Then
    '                    'If InStr(VALUE, "<BR>") > 0 Then
    '                    If InStr(VALUE, vbCrLf) > 0 Then
    '                        DB.AddSQL(op, SQL1, SQL2, "REMARK", VALUE, DBUTIL.FieldTypes.ftText)
    '                    Else
    '                        DB.AddSQL(op, SQL1, SQL2, "ANSWER", VALUE, DBUTIL.FieldTypes.ftText)
    '                    End If
    '                End If

    '                'DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", , DBUtil.FieldTypes.ftText)
    '                'DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", , DBUtil.FieldTypes.ftText)
    '                SQL = DB.CombineSQL(op, SQL1, SQL2, "SURVEY_DETAILS", "")
    '                DB.ExecSQL(SQL)
    '            End If
    '        Next



    '        SQL = "" : SQL1 = "" : SQL2 = ""

    '        ' insert user
    '        DB.AddSQL(op, SQL1, SQL2, "SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "FORM_ID", LookupSQL("SELECT FORM_ID FROM BATCHES WHERE BATCH_ID = " + BatchID), DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "SURVEY_DATE", Now, DBUTIL.FieldTypes.ftDate)
    '        DB.AddSQL(op, SQL1, SQL2, "SOURCE_ID", "Web", DBUTIL.FieldTypes.ftText)
    '        'SCORE 
    '        'DB.AddSQL(op, SQL1, SQL2, "CODE", txtCode.Text, DBUtil.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "NAME", Name, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "SURNAME", SurName, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "SEX", Sex, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "AGE", Age, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "EDUCATION", Education, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "WORK_TYPE", WorkType, DBUTIL.FieldTypes.ftNumeric)
    '        DB.AddSQL(op, SQL1, SQL2, "ADDRESS1", Address1, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "ADDRESS2", Address2, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "ADDRESS3", Address3, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "ADDRESS4", Address4, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "TUMBON", Tumbon, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "AMPHUR", Amphur, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "PROVINCE", Province, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "ZIPCODE", ZipCode, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "COMPANY_NAME", CompanyName, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "WORK_TYPE_OTHER", WorkTypeOther, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "EMAIL", EMail, DBUTIL.FieldTypes.ftText)
    '        DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDate)
    '        'DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", AppDateValue(System.DateTime.Now), CellDateTime)

    '        'If txtAge.Text <> "" And IsNumeric(txtAge.Text) Then
    '        '    If CInt(txtAge.Text) > 0 Then
    '        '        DB.AddSQL(op, SQL1, SQL2, "AGE", txtAge.Text, DBUtil.FieldTypes.ftText)
    '        '    End If
    '        'End If

    '        SQL = DB.CombineSQL(op, SQL1, SQL2, "SURVEYS", "")
    '        'SQL = CombineSQL2(op, SQL1, SQL2, "SURVEYS", "")
    '        DB.ExecSQL(SQL)
    '        Return ""
    '        'Return SQL

    '    Catch ex As Exception
    '        Return "เกิดข้อผิดพลาด : " & ex.Message
    '        'Return SQL
    '    End Try
    'End Function

    Public Function SaveSurvey(ByRef SurveyID As String, ByRef BatchID As String _
    , ByVal Name As String, ByVal SurName As String, ByVal CompanyName As String _
    , ByVal TelNo As String, ByVal EMail As String, ByVal PositionName As String, ByVal Request As HttpRequest _
    , ByVal controlPassed5 As String(), ByVal controlPassed6 As String) As String
        Dim SQL1, SQL2, SQL As String
        Dim MaxID As Long = 0
        Dim SysFileName As String = ""
        Dim op As Integer
        Dim op2 As Integer
        Dim Obj As Object
        Dim VALUE, controlPassed, tmpSQL As String
        Dim QID, CID, tmpCID, i, r As Integer
        Dim statusQ5, statusT5 As Boolean
        'Dim controlPassed5(), controlPassed6 As String

        Try
            If SurveyID <> "0" Then
                DB.ExecSQL("DELETE FROM SURVEYS WHERE SURVEY_ID=" & SurveyID)
                DB.ExecSQL("DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID=" & SurveyID)
                op2 = opUPDATE
            Else
                SurveyID = ToInt("0" & LookupSQL("SELECT MAX(SURVEY_ID) FROM SURVEYS")) + 1
                op2 = opINSERT
            End If


            op = opINSERT
            r = 0
            controlPassed = ""
            For Each Obj In Request.Form

                SQL = "" : SQL1 = "" : SQL2 = "" : VALUE = ""
                tmpSQL = "" : tmpCID = -1
                QID = -1 : CID = -1
                statusT5 = True : statusQ5 = True ' ให้สามารถเข้าไปเก็บชื่อใน controlPassed5() ได้

                Select Case Mid(CStr(Obj), 1, 4)
                    Case "txt_"
                        If Request.Form(CStr(Obj)) & "" <> "" Then
                            QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
                            'CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                            If CStr(Obj).IndexOf("C") = CStr(Obj).Length - 1 Then
                                CID = 0
                            Else
                                CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                            End If
                            VALUE = ValidateData(Request.Form(CStr(Obj)) & "") 'Replace(Request.Form(CStr(Obj)) & "", vbCrLf, "<BR>")
                        End If
                    Case "rdb_"
                        QID = CInt(Mid(CStr(Obj), 6))
                        CID = CInt(Mid(Request.Form(CStr(Obj)), Request.Form(CStr(Obj)).IndexOf("C") + 2))
                    Case "chk_"
                        QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
                        CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                    Case "rdt_" ' เป็น radio ที่มี textbox ด้วย 
                        Select Case Mid(CStr(Obj), 1, 5)
                            Case "rdt_T" ' ถ้ามี textbox แสดงว่าต้องมีการติ๊กด้วย
                                If Request.Form(CStr(Obj)) & "" <> "" Then ' user กรอกค่าใน txtbox
                                    QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
                                    CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                                    'QID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("Q") + 2, CStr(Request.Form(CStr(Obj))).IndexOf("C") - CStr(Request.Form(CStr(Obj))).IndexOf("Q") - 1))
                                    'CID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("C") + 2))
                                    VALUE += ValidateData(Request.Form(CStr(Obj)))
                                    For i = 0 To controlPassed5.Length - 1
                                        If InStr(controlPassed5(i), "rdt_Q" & QID) > 0 Then ' ถ้ามีการ save chkbox ไปก่อนแล้ว
                                            tmpCID = CInt(Mid(controlPassed5(i), controlPassed5(i).IndexOf("C") + 2))
                                            tmpSQL = "DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID = " & SurveyID & " and QUESTION_ID = " & QID & " AND CHOICE_ID = " & tmpCID
                                            DB.ExecSQL(tmpSQL)
                                            controlPassed5(i) = "chk_Q" & QID & "C" & CID & ""
                                            statusT5 = False
                                        End If
                                    Next
                                    If statusT5 Then
                                        controlPassed5(r) = "chk_Q" & QID & "C" & CID & ""
                                        r += 1
                                    End If
                                End If

                            Case "rdt_Q"
                                For i = 0 To controlPassed5.Length - 1
                                    If InStr(controlPassed5(i), "Q" & CInt(Mid(CStr(Obj), 6))) <> 0 Then ' control ถูก Save ลงไน DB แล้ว
                                        statusQ5 = False ' เคย save ค่านี้แล้ว
                                    End If
                                Next
                                If statusQ5 Then
                                    QID = CInt(Mid(CStr(Obj), 6))
                                    CID = CInt(Mid(Request.Form(CStr(Obj)), Request.Form(CStr(Obj)).IndexOf("C") + 2))

                                    controlPassed5(r) = Request.Form(CStr(Obj))
                                    r += 1
                                End If
                        End Select
                    Case "cht_"
                        Select Case Mid(CStr(Obj), 1, 5)
                            Case "cht_T" ' ถ้ามี textbox แสดงว่าต้องมีการติ๊กด้วย
                                If Request.Form(CStr(Obj)) & "" <> "" Then ' user กรอกค่าใน txtbox
                                    'QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
                                    'CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                                    QID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("Q") + 2, CStr(Request.Form(CStr(Obj))).IndexOf("C") - CStr(Request.Form(CStr(Obj))).IndexOf("Q") - 1))
                                    CID = CInt(Mid(Request.Form(CStr(Obj)) & "", CStr(Request.Form(CStr(Obj))).IndexOf("C") + 2))
                                    VALUE += ValidateData(Request.Form(CStr(Obj)))
                                    If InStr(controlPassed, "cht_Q" & QID & "C" & CID) > 0 Then ' ถ้ามีการ save chkbox ไปก่อนแล้ว
                                        tmpSQL = "DELETE FROM SURVEY_DETAILS WHERE SURVEY_ID = " & SurveyID & " and QUESTION_ID = " & QID & " AND CHOICE_ID = " & CID
                                        DB.ExecSQL(tmpSQL)
                                        controlPassed += "," & "chk_TQ" & QID & "C" & CID
                                    Else ' ยังไม่มีการ save control ที่ผ่านไปแล้ว
                                        controlPassed += "," & "chk_Q" & QID & "C" & CID & "," & "chk_TQ" & QID & "C" & CID
                                    End If
                                End If
                            Case "cht_Q" ' ถ้ามีการติ๊ก ต้องเช็กว่ามี textbox ด้วยรึเปล่า
                                If InStr(controlPassed, CStr(Obj)) = 0 Then ' control ยังไม่ถูก Save ลงไน DB
                                    QID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("Q") + 2, CStr(Obj).IndexOf("C") - CStr(Obj).IndexOf("Q") - 1))
                                    CID = CInt(Mid(CStr(Obj), CStr(Obj).IndexOf("C") + 2))
                                End If
                                controlPassed += "," & CStr(Obj)
                        End Select

                End Select
                If QID <> -1 And CID <> -1 Then
                    DB.AddSQL(op, SQL1, SQL2, "SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "QUESTION_ID", QID, DBUTIL.FieldTypes.ftNumeric)
                    DB.AddSQL(op, SQL1, SQL2, "CHOICE_ID", CID, DBUTIL.FieldTypes.ftNumeric)

                    If VALUE <> "" Then
                        'If InStr(VALUE, "<BR>") > 0 Then
                        If InStr(VALUE, vbCrLf) > 0 Then
                            DB.AddSQL(op, SQL1, SQL2, "REMARK", VALUE, DBUTIL.FieldTypes.ftText)
                        Else
                            DB.AddSQL(op, SQL1, SQL2, "ANSWER", VALUE, DBUTIL.FieldTypes.ftText)
                        End If
                    End If

                    'DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", , DBUtil.FieldTypes.ftText)
                    'DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", , DBUtil.FieldTypes.ftText)
                    SQL = DB.CombineSQL(op, SQL1, SQL2, "SURVEY_DETAILS", "")
                    DB.ExecSQL(SQL)
                End If
            Next



            SQL = "" : SQL1 = "" : SQL2 = ""

            ' insert user
            DB.AddSQL(op2, SQL1, SQL2, "SURVEY_ID", SurveyID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(op2, SQL1, SQL2, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(op2, SQL1, SQL2, "FORM_ID", LookupSQL("SELECT FORM_ID FROM BATCHES WHERE BATCH_ID = " + BatchID), DBUTIL.FieldTypes.ftNumeric)
            DB.AddSQL(op2, SQL1, SQL2, "SURVEY_DATE", Now, DBUTIL.FieldTypes.ftDate)
            'DB.AddSQL(op2, SQL1, SQL2, "SOURCE_ID", "Web", DBUTIL.FieldTypes.ftText)
            'SCORE 
            'DB.AddSQL(op, SQL1, SQL2, "CODE", txtCode.Text, DBUtil.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "NAME", Name, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "SURNAME", SurName, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "SEX", Sex, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "AGE", Age, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddSQL(op, SQL1, SQL2, "EDUCATION", Education, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddSQL(op, SQL1, SQL2, "WORK_TYPE", WorkType, DBUTIL.FieldTypes.ftNumeric)
            'DB.AddSQL(op, SQL1, SQL2, "ADDRESS1", Address1, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "ADDRESS2", Address2, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "ADDRESS3", Address3, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "ADDRESS4", Address4, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "TUMBON", Tumbon, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "AMPHUR", Amphur, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "PROVINCE", Province, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "ZIPCODE", ZipCode, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "TEL_NO", TelNo, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "COMPANY_NAME", CompanyName, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "POSITION_NAME", PositionName, DBUTIL.FieldTypes.ftText)
            'DB.AddSQL(op, SQL1, SQL2, "WORK_TYPE_OTHER", WorkTypeOther, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "EMAIL", EMail, DBUTIL.FieldTypes.ftText)
            DB.AddSQL(op2, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDate)
            'DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", AppDateValue(System.DateTime.Now), CellDateTime)

            'If txtAge.Text <> "" And IsNumeric(txtAge.Text) Then
            '    If CInt(txtAge.Text) > 0 Then
            '        DB.AddSQL(op, SQL1, SQL2, "AGE", txtAge.Text, DBUtil.FieldTypes.ftText)
            '    End If
            'End If

            SQL = DB.CombineSQL(op2, SQL1, SQL2, "SURVEYS", "")
            'SQL = CombineSQL2(op, SQL1, SQL2, "SURVEYS", "")
            DB.ExecSQL(SQL)
            Return ""
            'Return SQL

        Catch ex As Exception
            Return "เกิดข้อผิดพลาด : " & ex.Message
            'Return SQL
        End Try
    End Function

    Public Function SearchPersonType(Optional ByVal PersonType As String = "", Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            If PersonType & "" <> "" Then DB.AddCriteria(CriteriaSQL, "PERSON_TYPE", PersonType, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * " & _
                 "FROM VW_PERSON_TYPES "

            If CriteriaSQL <> "" Then
                SQL += " AND " + CriteriaSQL
            End If
            SQL &= " ORDER BY PERSON_TYPE "
            DB.OpenDT(DT, SQL)
            Return DT

        Catch ex As Exception
            Throw (ex)
        End Try
    End Function

    Public Function ManageBatch(ByVal op As Integer, ByVal BatchID As String, Optional ByVal BatchName As String = Nothing _
    , Optional ByVal FormID As String = Nothing, Optional ByVal StartDate As String = Nothing, Optional ByVal EndDate As String = Nothing _
    , Optional ByVal RegisFlag As String = Nothing, Optional ByVal UserType As String = Nothing _
    , Optional ByVal OtherCriteria As String = "") As String
        Dim SQL1, SQL2, SQL As String
        Dim Criteria As String = ""

        Try
            SQL = "" : SQL1 = "" : SQL2 = ""
            If op <> opINSERT Then
                Criteria = OtherCriteria
                DB.AddCriteria(Criteria, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            End If
            If op <> opDELETE Then
                If op = opINSERT Then
                    op = opINSERT
                    If Not IsNothing(BatchID) Then DB.AddSQL(op, SQL1, SQL2, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
                Else
                    op = opUPDATE
                End If
                If Not IsNothing(BatchName) Then DB.AddSQL(op, SQL1, SQL2, "BATCH_NAME", BatchName, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(FormID) Then DB.AddSQL(op, SQL1, SQL2, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)
                If Not IsNothing(StartDate) Then DB.AddSQL(op, SQL1, SQL2, "START_DATE", AppDateValue(StartDate), DBUTIL.FieldTypes.ftDate)
                If Not IsNothing(EndDate) Then DB.AddSQL(op, SQL1, SQL2, "END_DATE", AppDateValue(EndDate), DBUTIL.FieldTypes.ftDate)
                If Not IsNothing(RegisFlag) Then DB.AddSQL(op, SQL1, SQL2, "REGISTER_FLAG", RegisFlag, DBUTIL.FieldTypes.ftText)
                If Not IsNothing(UserType) Then DB.AddSQL(op, SQL1, SQL2, "USER_TYPE", UserType, DBUTIL.FieldTypes.ftNumeric)
                DB.AddSQL(op, SQL1, SQL2, "DATE_UPDATED", Now, DBUTIL.FieldTypes.ftDateTime)
                DB.AddSQL(op, SQL1, SQL2, "USER_UPDATED", HttpContext.Current.Session("USER_NAME") & "", DBUTIL.FieldTypes.ftText)
            End If

            If op <> opINSERT AndAlso Criteria = "" Then
                Throw New Exception("Insufficient data!")
            Else
                SQL = DB.CombineSQL(op, SQL1, SQL2, "BATCHES", Criteria)
                DB.ExecSQL(SQL)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function SearchAgeData(Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            'If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            'If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM REF_AGE_DATA"
            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            SQL &= " ORDER BY AGE_ID"
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchEducationData(Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            'If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            'If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM REF_EDUCATION_DATA"
            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            SQL &= " ORDER BY EDUCATION_ID"
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function SearchWorkType(Optional ByVal OtherCriteria As String = "") As DataTable
        Dim SQL As String
        Dim CriteriaSQL As String
        Dim DT As DataTable = Nothing

        Try
            CriteriaSQL = OtherCriteria
            'If BatchID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "BATCH_ID", BatchID, DBUTIL.FieldTypes.ftNumeric)
            'If FormID & "" <> "" Then DB.AddCriteria(CriteriaSQL, "FORM_ID", FormID, DBUTIL.FieldTypes.ftNumeric)

            SQL = "SELECT * FROM REF_WORK_TYPES"
            If CriteriaSQL <> "" Then
                SQL += " WHERE " + CriteriaSQL
            End If
            SQL &= " ORDER BY WORK_TYPE"
            DB.OpenDT(DT, SQL)
            Return DT
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
