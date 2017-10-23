<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey
	Dim orgChangID
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere, strFg
  
     On Error Resume Next    
    Err.Clear                                                               '☜: Clear Error status

    strWhere = strWhere & "org_change_dt  = ( select MAX(org_change_dt) "
    strWhere = strWhere & "					from b_acct_dept "
    strWhere = strWhere & "					where org_change_dt <= " & FilterVar(UNIConvDate(lgKeyStream(0)), "''", "S")
    strWhere = strWhere & ") "

    call CommonQueryRs(" max(org_change_id) "," b_acct_dept ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    orgChangID = Trim(Replace(lgF0,Chr(11),""))

    strWhere =  FilterVar(UNIConvDate(lgKeyStream(0)), "''", "S")
    strWhere = strWhere & " And HCA060T.Dilig_cd LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
    strWhere = strWhere & " And (B_ACCT_DEPT.internal_cd >= " & FilterVar(lgKeyStream(2), "''", "S")
    strWhere = strWhere & " And B_ACCT_DEPT.internal_cd <= " & FilterVar(lgKeyStream(3), "''", "S") & ")"
    strWhere = strWhere & " And HCA060T.EMP_NO LIKE " & FilterVar(lgKeyStream(4),"'%'", "S")
    strWhere = strWhere & " order by HCA060T.EMP_NO,HCA060T.Dilig_cd "
	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '☜ : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME") )                 

' 입력된일자의 근무당시의 부서코드검색 / 2002.04.08 송봉규 
' 부서정보 가져오는 로직처리 SQL에서 함/  2003.9.17 by lsn 	

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_HH"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
                       
        Loop 
    End If
           
      If iDx <= C_SHEETMAXROWS_D Then
         lgStrPrevKey = ""
      End If   

      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
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
	Dim iCnt, strFg, strType, strOrgId
	Dim strHoli_type, strHoliday_apply
	Dim f_gazet_dt,f_emp_no
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
  
	f_gazet_dt = FilterVar(UNIConvDate(arrColVal(2)),NULL,"S")
	f_emp_no = FilterVar(UCase(arrColVal(3)), "''", "S")
	
' 부서정보 MA에서 가져옴 /  2003.9.17 by lsn 
	Call CommonQueryRs(" cost_cd "," b_acct_dept ", " org_change_dt = (select MAX(org_change_dt) from b_acct_dept where org_change_dt <= " & _
	f_gazet_dt & ") and dept_cd =  " & FilterVar(arrColVal(8), "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Replace(lgF0, Chr(11), "")

	If IsNull(strFg) or strFg = ""  or strFg = "X" Then
		Call SubHandleError("MS",lgObjConn,lgObjRs,Err)
		Exit Sub
	End If
   	
    Call CommonQueryRs(" biz_area_cd "," b_cost_center ", " cost_cd =  " & FilterVar(strFg , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strOrgId = Replace(lgF0, Chr(11), "")
  
    Call CommonQueryRs(" wk_type "," HCA040T "," emp_no = " & f_emp_no & _
    " and chang_dt = (select max(chang_dt) from hca040t where emp_no = " & f_emp_no & _
    " and chang_dt <= " & f_gazet_dt & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strType = Replace(lgF0, Chr(11), "")
	
	If (strType = "X" OR strType = "") Then strType = "0"

' 해당일이 휴일인지 평일인지 가져옴 2002.11.06 by sbk
	Call CommonQueryRs(" holi_type "," HCA020T "," org_cd =  " & FilterVar(strOrgId , "''", "S") & "" & _
                       " and wk_type =  " & FilterVar(strType , "''", "S") & " and date = " & f_gazet_dt ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strFg = Replace(lgF0, Chr(11), "")
	
	If IsNull(strFg) or strFg = ""  or strFg = "X" Then
		Call SubHandleError("MY",lgObjConn,lgObjRs,Err)
		Exit Sub
	else
		strHoli_type = Replace(lgF0, Chr(11), "")		
	End If
	
    Call CommonQueryRs(" day_time, holiday_apply "," HCA010T "," dilig_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strFg = Replace(lgF0, Chr(11), "") '근태코드에서 일수/시간 판별 
	strHoliday_apply = Replace(lgF1, Chr(11), "")

'2002.11.06 by sbk 휴일적용여부가 'N'이고 해당일이 휴일이면 등록 불가 
	If strHoliday_apply = "N" AND strHoli_type = "H" Then 
        Call DisplayMsgBox("800505", vbInformation, Trim(arrColVal(7)), "", I_MKSCRIPT)           
        ObjectContext.SetAbort
        Call SetErrorStatus
    	Exit Sub        
	End If	
'2003.09.04 by lsn 일수근태와 시간근태 동시에 저장 가능 
    If strFg = "1" Then   '일수근태인경우 
 		Call CommonQueryRs(" count(*) "," HCA060T "," emp_no = " & f_emp_no & _
			" and dilig_dt = " & f_gazet_dt & " and dilig_cd in (SELECT dilig_cd from hca010t where day_time = " & FilterVar("1", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Else  '시간근태인경우 
	    Call CommonQueryRs(" count(*) "," HCA060T "," emp_no = " & f_emp_no & _
	    " and dilig_dt = " & f_gazet_dt & _
	    " and dilig_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S") & _
	    " and dilig_cd in (SELECT dilig_cd from hca010t where day_time=" & FilterVar("2", "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
  	End If 	

	iCnt = Replace(lgF0, Chr(11), "")	 

	If iCnt > 0 Then 	
		Call SubHandleError("MX",lgObjConn,lgObjRs,Err)
		Exit Sub
	Else
		lgStrSQL = "INSERT INTO HCA060T("
		lgStrSQL = lgStrSQL & " EMP_NO, " 
		lgStrSQL = lgStrSQL & " DILIG_DT, " 
		lgStrSQL = lgStrSQL & " DILIG_CD, " 
		lgStrSQL = lgStrSQL & " DILIG_CNT, " 
		lgStrSQL = lgStrSQL & " DILIG_HH, " 
		lgStrSQL = lgStrSQL & " DILIG_MM, "
		lgStrSQL = lgStrSQL & " ISRT_DT ," 
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO ," 
		lgStrSQL = lgStrSQL & " UPDT_DT ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
    
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S")		& ","        
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & 1												    & ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)                        & ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)                        & ","    
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")
		lgStrSQL = lgStrSQL & ")"  
    
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
		If strFg = "1" Then
			lgStrSQL = "INSERT INTO HCA050T("
			lgStrSQL = lgStrSQL & " EMP_NO, " 
			lgStrSQL = lgStrSQL & " DILIG_CD, " 
			lgStrSQL = lgStrSQL & " DILIG_STRT_DT, " 
			lgStrSQL = lgStrSQL & " DILIG_END_DT, " 
			lgStrSQL = lgStrSQL & " REMARK, "
			lgStrSQL = lgStrSQL & " ISRT_DT ," 
			lgStrSQL = lgStrSQL & " ISRT_EMP_NO ," 
			lgStrSQL = lgStrSQL & " UPDT_DT ," 
			lgStrSQL = lgStrSQL & " UPDT_EMP_NO )" 
			lgStrSQL = lgStrSQL & " VALUES(" 
    
			lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","		        
			lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
			lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S")		& ","
			lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S")		& ","
			lgStrSQL = lgStrSQL & "'',"
            lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & ","
			lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
            lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & ","
			lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        
			lgStrSQL = lgStrSQL & ")"      
    
			lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
			Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		End If
	End If
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
    lgStrSQL = "UPDATE  HCA060T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_HH =       " & UNIConvNum(arrColVal(5),0)                        & ","
    lgStrSQL = lgStrSQL & " DILIG_MM =       " & UNIConvNum(arrColVal(6),0)                        & ","
    lgStrSQL = lgStrSQL & " UPDT_DT  =       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO =    " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   =       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_DT =       " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_CD =       " & FilterVar(UCase(arrColVal(4)), "''", "S")
  
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_HH =       " & UNIConvNum(arrColVal(5),0)                        & ","
    lgStrSQL = lgStrSQL & " DILIG_MM =       " & UNIConvNum(arrColVal(6),0)                        & ","
    lgStrSQL = lgStrSQL & " UPDT_DT  =       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO =    " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   =       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_STRT_DT =	 " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_CD =       " & FilterVar(UCase(arrColVal(4)), "''", "S")
  
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

    lgStrSQL = "DELETE  HCA060T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   =   " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_DT =   " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_CD =   " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
	lgStrSQL = "DELETE  HCA050T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   =   " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_CD =   " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "    
    lgStrSQL = lgStrSQL & " DILIG_STRT_DT = " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_END_DT  = " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	lgStrSQL = "DELETE  E11070T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   =   " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " DILIG_CD =   " & FilterVar(UCase(arrColVal(4)), "''", "S") & " AND "    
    lgStrSQL = lgStrSQL & " DILIG_STRT_DT = " & FilterVar(UNIConvDate(arrColVal(2)),NULL,"S") 
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
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
                       lgStrSQL = "Select TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " HCA060T.EMP_NO, "   
                       lgStrSQL = lgStrSQL & " HAA010T.NAME, "  
                       lgStrSQL = lgStrSQL & " z.DEPT_CD, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetDeptName(z.DEPT_CD,HCA060T.DILIG_DT) DEPT_NM,"
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0002', HAA010T.ROLL_PSTN) ROLL_PSTN_NM, "
                       lgStrSQL = lgStrSQL & " HCA060T.DILIG_CD, "
                       lgStrSQL = lgStrSQL & " HCA010T.DILIG_NM, HCA010T.DAY_TIME, " 
                       lgStrSQL = lgStrSQL & " HCA060T.DILIG_HH, HCA060T.DILIG_MM "

                       lgStrSQL = lgStrSQL & " From  HCA060T, HAA010T, HCA010T, B_ACCT_DEPT ,"
                       lgStrSQL = lgStrSQL & "     ( select x.emp_no,max(y.dept_cd) dept_cd"
                       lgStrSQL = lgStrSQL & "		from (	select a.emp_no, max(a.gazet_dt) gazet_dt "
                       lgStrSQL = lgStrSQL & "				from hba010t a "
                       lgStrSQL = lgStrSQL & "				where a.gazet_dt  <= " &  FilterVar(UNIConvDate(lgKeyStream(0)), "''", "S")
                       lgStrSQL = lgStrSQL & "				group by a.emp_no) x left outer join hba010t y on x.emp_no=y.emp_no and x.gazet_dt = y.gazet_dt "
                       lgStrSQL = lgStrSQL & "		group by x.emp_no, x.gazet_dt "
                       lgStrSQL = lgStrSQL & "		) z "
                       lgStrSQL = lgStrSQL & " Where HAA010T.EMP_NO = HCA060T.EMP_NO " 
                       lgStrSQL = lgStrSQL & "   and B_ACCT_DEPT.ORG_CHANGE_ID =  " & FilterVar(orgChangID , "''", "S") 
                       lgStrSQL = lgStrSQL & "   and HCA060T.EMP_NO = z.EMP_NO " 
                       lgStrSQL = lgStrSQL & "   and B_ACCT_DEPT.DEPT_CD = z.DEPT_CD " 
                       lgStrSQL = lgStrSQL & "   and HCA060T.DILIG_CD = HCA010T.DILIG_CD and HCA060T.DILIG_DT = " & pCode 
          End Select
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
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
        Case "MS"
                 Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
        Case "MV"
                 Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MZ"
                 Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
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
    End Select       
       
</Script>	
