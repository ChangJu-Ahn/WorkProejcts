<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
	Const C_SHEETMAXROWS_D  = 100
    '---------------------------------------Common-----------------------------------------------------------
    Dim txtPayAmt

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             CALL SubBizSaveMultiDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim txtGlNo
    Dim iLcNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    Dim txtBonusCd
    Dim txtBonusNm
    Dim txtBizAreaCd
    Dim txtBizAreaNm
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Call CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & "  AND MINOR_TYPE=" & FilterVar("U", "''", "S") & "  and minor_cd = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if Trim(Replace(lgF0,Chr(11),"")) = "X" then
        txtBonusCd = ""
        txtBonusNm = ""
        Call DisplayMsgBox("800142", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
    else
        txtBonusCd =  Trim(lgKeyStream(1))
        txtBonusNm = Trim(Replace(lgF0,Chr(11),""))
        If lgKeyStream(2) <> "" Then
            Call CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ", " BIZ_AREA_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            if Trim(Replace(lgF0,Chr(11),"")) = "X" then
                txtBizAreaCd = ""
                txtBizAreaNm = ""
                Call DisplayMsgBox("800142", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
            else
                txtBizAreaCd =  Trim(lgKeyStream(2))
                txtBizAreaNm = Trim(Replace(lgF0,Chr(11),""))
            End If
        End If
    end if
    Call SubMakeSQLStatements("MR","X","X",C_EQ)                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()

    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("org_change_id"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("internal_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("acct_type"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("pay_amt"), ggQty.DecPoint, 0)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)


		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
        Call SubMakeSQLStatements("MT",strWhere,"X",C_EQ)                                 '☆ : Make sql statements

        If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
            lgStrPrevKeyIndex = ""
		    Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
            Call SetErrorStatus()
        Else
            txtPayAmt = UNINumClientFormat(lgObjRs("pay_amt"), ggQty.DecPoint, 0)
        End If
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtBonusCd.value = "<%=ConvSPChars(txtBonusCd)%>"
		.frm1.txtBonus.value = "<%=ConvSPChars(txtBonusNm)%>"
		.frm1.txtBizAreaCd.value = "<%=ConvSPChars(txtBizAreaCd)%>"
		.frm1.txtBizArea.value = "<%=ConvSPChars(txtBizAreaNm)%>"
	END With
</SCRIPT>
<%
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select

        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

	'데이터 중복 체크 
	lgStrSQL =			  " select  count(PAY_TYPE) as Cnt "
	lgStrSQL = lgStrSQL & " from	a_paid_year_bonus			  "
	lgStrSQL = lgStrSQL & " where	YYYY	=	"	& FilterVar(lgKeyStream(0), "''", "S")
	lgStrSQL = lgStrSQL & " and		PAY_TYPE =	"	& FilterVar(UCase(lgKeyStream(1)), "''", "S")
	lgStrSQL = lgStrSQL & " and		DEPT_CD =	"	& FilterVar(arrColVal(2), "''", "S")
	lgStrSQL = lgStrSQL & " and		BIZ_AREA_CD =	"	& FilterVar(arrColVal(3), "''", "S")
	lgStrSQL = lgStrSQL & " and		ORG_CHANGE_ID =	"	& FilterVar(arrColVal(4), "''", "S")
	lgStrSQL = lgStrSQL & " and		INTERNAL_CD =	"	& FilterVar(arrColVal(5), "''", "S")
	lgStrSQL = lgStrSQL & " and		ACCT_TYPE =	"	& FilterVar(arrColVal(6), "''", "S")
		
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		If cdbl(lgObjRs("Cnt")) <> 0  Then
			Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)
			Call SetErrorStatus
			Exit Sub
	   End If
	End if
	
    lgStrSQL = "INSERT INTO a_paid_year_bonus ("
    lgStrSQL = lgStrSQL & " YYYY        ,"
    lgStrSQL = lgStrSQL & " PAY_TYPE    ,"
    lgStrSQL = lgStrSQL & " DEPT_CD     ,"
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD ,"
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID ,"
    lgStrSQL = lgStrSQL & " INTERNAL_CD ,"
    lgStrSQL = lgStrSQL & " ACCT_TYPE   ,"
    lgStrSQL = lgStrSQL & " PAY_AMT         ,"
    lgStrSQL = lgStrSQL & " PAY_MM         ,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID ,"
    lgStrSQL = lgStrSQL & " INSRT_DT	  ,"
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,"
    lgStrSQL = lgStrSQL & " UPDT_DT    )"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(lgKeyStream(1)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

	Response.Write lgStrSQL
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  a_paid_year_bonus "
    lgStrSQL = lgStrSQL & " SET   DEPT_CD = " & FilterVar(UCase(arrColVal(2)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & "       BIZ_AREA_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & "       ORG_CHANGE_ID = " & FilterVar(UCase(arrColVal(4)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & "       INTERNAL_CD = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & "       ACCT_TYPE = " & FilterVar(UCase(arrColVal(6)), "''", "S") & ", "
    lgStrSQL = lgStrSQL & "       PAY_AMT = " & UNIConvNum(arrColVal(7),0)
    lgStrSQL = lgStrSQL & " WHERE  YYYY = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "  AND   PAY_TYPE = " & FilterVar(UCase(lgKeyStream(1)), "''", "S")
    lgStrSQL = lgStrSQL & "  AND   DEPT_CD = " & FilterVar(UCase(arrColVal(8)), "''", "S")
    lgStrSQL = lgStrSQL & "  AND   ACCT_TYPE = " & FilterVar(UCase(arrColVal(9)), "''", "S")
    lgStrSQL = lgStrSQL & "  AND   PAY_MM  = " & FilterVar(lgKeyStream(3), "''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  a_paid_year_bonus "
    lgStrSQL = lgStrSQL & " WHERE  YYYY = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")
    lgStrSQL = lgStrSQL & "   and  PAY_TYPE      = " & FilterVar(UCase(lgKeyStream(1)), "''", "S")
    lgStrSQL = lgStrSQL & "   and  DEPT_CD       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   and  ACCT_TYPE     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND  PAY_MM  = " & FilterVar(lgKeyStream(3), "''", "S")



    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1

            Select Case Mid(pDataType,2,1)
                Case "R"
                    lgStrSQL = "Select TOP " & iSelCount  & " a.dept_cd, "
                    lgStrSQL = lgStrSQL & "                   c.dept_nm, "
                    lgStrSQL = lgStrSQL & "                   a.acct_type, "
                    lgStrSQL = lgStrSQL & "                   b.minor_nm, "
                    lgStrSQL = lgStrSQL & "                   a.biz_area_cd, "
                    lgStrSQL = lgStrSQL & "                   a.org_change_id, "
                    lgStrSQL = lgStrSQL & "                   a.internal_cd, "
                    lgStrSQL = lgStrSQL & "                   isnull(a.pay_amt,0) as pay_amt "
                    lgStrSQL = lgStrSQL & "             from  a_paid_year_bonus a, "
                    lgStrSQL = lgStrSQL & "                   b_minor b, "
                    lgStrSQL = lgStrSQL & "                   b_acct_dept c "
                    lgStrSQL = lgStrSQL & "            where  a.acct_type = b.minor_cd "
                    lgStrSQL = lgStrSQL & "              and  b.major_cd = " & FilterVar("H0071", "''", "S") & "  "
                    lgStrSQL = lgStrSQL & "              and  a.dept_cd = c.dept_cd "
                    lgStrSQL = lgStrSQL & "              and  a.org_change_id = c.org_change_id "
                    lgStrSQL = lgStrSQL & "              and  a.yyyy = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & "              and  a.pay_type = " & FilterVar(lgKeyStream(1), "''", "S")
                    lgStrSQL = lgStrSQL & "              and  a.pay_mm = " & FilterVar(lgKeyStream(3), "''", "S")
                    lgStrSQL = lgStrSQL & "              and  a.biz_area_cd = " & FilterVar(lgKeyStream(2), "''", "S")
                    lgStrSQL = lgStrSQL & "         group by  a.dept_cd, "
                    lgStrSQL = lgStrSQL & "                   c.dept_nm, "
                    lgStrSQL = lgStrSQL & "                   a.acct_type, "
                    lgStrSQL = lgStrSQL & "                   b.minor_nm, "
                    lgStrSQL = lgStrSQL & "                   a.biz_area_cd, "
                    lgStrSQL = lgStrSQL & "                   a.org_change_id, "
                    lgStrSQL = lgStrSQL & "                   a.internal_Cd, "
                    lgStrSQL = lgStrSQL & "                   isnull(a.pay_amt,0) "
                Case "T"
                    lgStrSQL = "Select                    isnull(sum(a.pay_amt),0) as pay_amt "
                    lgStrSQL = lgStrSQL & "         from  a_paid_year_bonus a, "
                    lgStrSQL = lgStrSQL & "               b_minor b, "
                    lgStrSQL = lgStrSQL & "               b_acct_dept c "
                    lgStrSQL = lgStrSQL & "        where  a.acct_type = b.minor_cd "
                    lgStrSQL = lgStrSQL & "          and  b.major_cd = " & FilterVar("H0071", "''", "S") & "  "
                    lgStrSQL = lgStrSQL & "          and  a.dept_cd = c.dept_cd "
                    lgStrSQL = lgStrSQL & "          and  a.org_change_id = c.org_change_id "
                    lgStrSQL = lgStrSQL & "          and  a.yyyy = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & "          and  a.pay_type = " & FilterVar(lgKeyStream(1), "''", "S")
                    lgStrSQL = lgStrSQL & "          and  a.pay_mm = " & FilterVar(lgKeyStream(3), "''", "S")
                    lgStrSQL = lgStrSQL & "          and  a.biz_area_cd = " & FilterVar(lgKeyStream(2), "''", "S")
           End Select
    End Select
  
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .frm1.txtPayAmt.text = "<%=txtPayAmt%>"
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
       Case "<%=UID_M0006%>"                                                         '☜ : Batch
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.ExeReflectOk
          Else
             Parent.ExeReflectNo
          End If
    End Select


</Script>
