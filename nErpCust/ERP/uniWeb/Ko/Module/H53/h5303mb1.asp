<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgGetSvrDateTime, strTab
        
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")
    lgGetSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
	strTab			  = Request("txtTab")                                           '¢Ð: Read Operation Mode (CRUD)

    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)        
             Call SubAutoQueryMulti()              
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	lgStrSQL = "SELECT NUM, BIZ_AREA_NUM, SUB_NUM, HOIGEI, UNIT_BIZ_AREA, CERTI_NUM, NAME, RES_NO, ACQ_DT, LAST_PAY_MONTH, LAST_PAY_TOT, LAST_BOSU_TOT, DUTY_MONTH "
	lgStrSQL = lgStrSQL & " FROM HDB030T "
	lgStrSQL = lgStrSQL & " WHERE  DIV=" & FilterVar(strTab, "''", "S") & " AND YEAR_YY =" & FilterVar(lgKeyStream(0), "''", "S") & " AND BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S")
	lgStrSQL = lgStrSQL & " ORDER BY NUM "
'Response.Write lgStrSQL
'Response.End
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HOIGEI"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UNIT_BIZ_AREA"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CERTI_NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RES_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACQ_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LAST_PAY_MONTH"))
'            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LAST_PAY_TOT"))
            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LAST_BOSU_TOT"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LAST_PAY_TOT"), ggAmtOfMoney.DecPoint, 0)  
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LAST_BOSU_TOT"), ggAmtOfMoney.DecPoint, 0)           
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUTY_MONTH"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LAST_PAY_MONTH"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1

        Loop 
    End If
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet
 
End Sub    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubAutoQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strSect_cd
    Dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    strSect_cd         = FilterVar(lgKeyStream(1), "''", "S")

    strWhere = iKey1
    strWhere = strWhere & " And a.year_area_cd = " & strSect_cd
    strWhere = strWhere & " And convert(varchar(8), isnull(b.med_acq_dt, a.entr_dt), 112) < " & iKey1 & " + '1231' "
    strWhere = strWhere & " And (b.med_loss_dt is null or convert(varchar(8), b.med_loss_dt, 112) > " & iKey1 & " + '1231') "
    strWhere = strWhere & " And a.emp_no = b.emp_no And b.emp_no = d.emp_no And d.emp_no *= t.emp_no And c.emp_no =* d.emp_no "
    strWhere = strWhere & " And c.year_yy =* d.year_yy "
    strWhere = strWhere & " And ( a.internal_cd  LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " ) "
    strWhere = strWhere & " Group By a.name, a.emp_no, b.med_insur_no, a.res_no, b.med_acq_dt, a.entr_dt, t.sub_tot_amt,d.income_tot_amt, d.non_tax5 "
    strWhere = strWhere & " Order By b.med_acq_dt, b.med_insur_no "

  
    Call SubMakeSQLStatements("MR",strWhere,iKey1,C_EQ)                              '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx        
            lgstrData = lgstrData & Chr(11) & lgKeyStream(3)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(4)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(5)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(6)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("med_insur_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("med_acq_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Pre_Prov_mon"))
'            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sub_tot_amt"))
 '           lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("income_tot_amt"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("sub_tot_amt"), ggAmtOfMoney.DecPoint, 0)  
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("income_tot_amt"), ggAmtOfMoney.DecPoint, 0)  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("work_month_amt"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1
        Loop 
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet
 
End Sub   
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data

    arrColVal = Split(arrRowVal(0), gColSep)   

    
    If  arrColVal(0) = "C" Then
		Call SubBizSaveMultiDelete(arrColVal)        
	End If

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data

        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
'            Case "D"
'                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

'strTab

    lgStrSQL = "INSERT INTO HDB030T("
    lgStrSQL = lgStrSQL & " DIV,"     
    lgStrSQL = lgStrSQL & " YEAR_YY," 
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD," 
    lgStrSQL = lgStrSQL & " NUM," 
    lgStrSQL = lgStrSQL & " BIZ_AREA_NUM,"
    lgStrSQL = lgStrSQL & " SUB_NUM," 
    lgStrSQL = lgStrSQL & " HOIGEI,"
    lgStrSQL = lgStrSQL & " UNIT_BIZ_AREA,"
    lgStrSQL = lgStrSQL & " CERTI_NUM,"
    lgStrSQL = lgStrSQL & " NAME,"
    lgStrSQL = lgStrSQL & " RES_NO,"
    lgStrSQL = lgStrSQL & " ACQ_DT,"
    lgStrSQL = lgStrSQL & " LAST_PAY_MONTH,"
    lgStrSQL = lgStrSQL & " LAST_PAY_TOT,"
    lgStrSQL = lgStrSQL & " LAST_BOSU_TOT,"
    lgStrSQL = lgStrSQL & " DUTY_MONTH," 
    
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(strTab, "''", "S")     & ","    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(12)), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S")    & ","
'    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(14)), "''", "S")   & ","
'    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(15)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0)					& ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(15),0)					& ","
    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(16)), "''", "S")    & ","
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"


'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,C_COMP_CD,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  HDB030T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ACQ_DT				= "     & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " LAST_PAY_MONTH		= "     & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " LAST_PAY_TOT		= "     & UNIConvNum(arrColVal(7),0)					& ","
    lgStrSQL = lgStrSQL & " LAST_BOSU_TOT		= "     & UNIConvNum(arrColVal(8),0)					& ","
    lgStrSQL = lgStrSQL & " DUTY_MONTH			= "     & FilterVar(UCase(arrColVal(9)), "''", "S")       
    
    lgStrSQL = lgStrSQL & " WHERE YEAR_YY		= " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND BIZ_AREA_CD		= " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND RES_NO			= " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DIV				= " & FilterVar(strTab, "''", "S")
'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,C_COMP_CD,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	lgStrSQL = "DELETE  HDB030T"
    lgStrSQL = lgStrSQL & " WHERE YEAR_YY			= " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "		AND BIZ_AREA_CD		= " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "		AND DIV				= " & FilterVar(strTab, "''", "S")
	lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MD",lgObjConn,C_COMP_CD,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                    lgStrSQL = "Select "
                    lgStrSQL = lgStrSQL & " a.name  name, a.emp_no  emp_no, left(replace(med_insur_no,'-',''),8) med_insur_no, "
                    lgStrSQL = lgStrSQL & " a.res_no,"
                    lgStrSQL = lgStrSQL & " convert(varchar(8),ISNULL(b.med_acq_dt, a.entr_dt),112) med_acq_dt,"
                    lgStrSQL = lgStrSQL & " ISNULL(t.sub_tot_amt, 0)  sub_tot_amt, "
                    lgStrSQL = lgStrSQL & " d.income_tot_amt + d.non_tax5 - ISNULL(SUM(ISNULL(c.a_pay_tot_amt,0)),0) - ISNULL(SUM(ISNULL(c.a_bonus_tot_amt,0)),0)  - ISNULL(SUM(ISNULL(c.a_after_bonus_amt,0)),0)   income_tot_amt, "
                    lgStrSQL = lgStrSQL & " CASE WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, a.entr_dt), 112) < " & pCode1 & " + '0101' "
                    lgStrSQL = lgStrSQL & " THEN DATEDIFF(month, CONVERT(DATETIME, " & pCode1 & " + '0101'), CONVERT(DATETIME, " & pCode1 & " + '1231')) + 1 "
                    lgStrSQL = lgStrSQL & " WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, a.entr_dt), 112) >= " & pCode1 & " + '0101' "
                    lgStrSQL = lgStrSQL & " THEN DATEDIFF(month, ISNULL(b.med_acq_dt, a.entr_dt), CONVERT(DATETIME, " & pCode1 & " + '1231')) + 1 END  work_month_amt,"
                    lgStrSQL = lgStrSQL & "(Select count(*) from hdf070t where emp_no = a.emp_no and PAY_YYMM LIKE " & pCode1 & " + '%' and med_insur > 0 and prov_type = '1') Pre_Prov_mon "                       
                    lgStrSQL = lgStrSQL & " From  haa010t a, hdf020t b, hfa040t c, hfa050t d, "
                    lgStrSQL = lgStrSQL & " (SELECT emp_no, SUM(sub_amt) AS sub_tot_amt FROM hdf060t WHERE sub_yymm LIKE " & pCode1 & " + '%' AND sub_cd in ('S01','SP1') and sub_type not in ('B','C','P','Q') GROUP BY emp_no) AS t "
                    lgStrSQL = lgStrSQL & " Where d.year_yy " & pComp & pCode
'Response.Write lgStrSQL
'Response.End                       
           End Select
    End Select
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk        
	         End with        
          Else   
          End If   
    End Select    
    
       
</Script>	

