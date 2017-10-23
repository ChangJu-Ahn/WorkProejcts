<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
    Dim lgGetSvrDateTime	
	Const C_SHEETMAXROWS_D = 100
    Call LoadBasisGlobalInf()
    Call loadInfTB19029B("Q", "H", "NOCOOKIE", "MB")
    lgGetSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
        
    Dim strFilePath
	Dim Pinfo,Fnm,CFnm,Pnm,FPnm      
    Dim Fso,DFnm,CTFnm
    Dim xdn

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection    
     
    Select Case lgOpModeCRUD
		Case CStr(UID_M0001)                                                         '☜: Query
			Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()        
        Case CStr(UID_M0003)                                                         '☜: Save,Update
            Set Fso = CreateObject("Scripting.FileSystemObject")  

            Pinfo = Request.ServerVariables ("PATH_INFO")                                       '현재File의 경로를 받는다.
            Fnm = Mid(Pinfo,InstrRev(Pinfo,"/")+1,InstrRev(Pinfo,".")-InstrRev(Pinfo,"/")-1)    'File의 경로중 File Name만 저장 
            Pnm = Mid(Pinfo,1,InstrRev(Pinfo,"/")+1)                                            'File Name 부분을 뺀 나머지 경로를 저장 
            FPnm = Server.MapPath("../../files/u2000/" & Fnm)			            '경로를 System 디렉토리로 바꾼다.

            Set CTFnm = Fso.CreateTextFile (Fpnm,true)		                                        'text를 저장할 File을 생성 

            Call SubBizFile()

            CTFnm.Write lgstrData                                                                'Text 내용부분           
            
            DFnm = Fso.GetFileName(FPnm)            
            CTFnm.close    
            Set CTFnm = nothing
            Set Fso = nothing           
' Response.Write "lgstrData:"  & lgstrData        
' Response.End             
            Call HideStatusWnd           
            
%>
    <SCRIPT LANGUAGE=VBSCRIPT>
				parent.subVatDiskOK("<%=DFnm%>")
	</SCRIPT>
<%
        Case  "7"                                                         '☜: Client로 File Copy
			Err.Clear 

			Call HideStatusWnd

			strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" _
				   & Request.ServerVariables("SERVER_PORT")
			If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
				strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
			End If
			strFilePath = strFilePath & "files/u2000/"
			strFilePath = strFilePath & Request("txtFileName")
'Response.Write "strFilePath:" & strFilePath
'Response.End			
        Case "5"
			Call SubAutoQuery()        
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = "SELECT BIZ_AREA_NUM, NUM, BIZ_PAGE, RES_NO, NAME, MONTH_CNT, PAY_TOT, SUB_CODE, EDI_CODE"
	lgStrSQL = lgStrSQL & " FROM HDB040T "
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY =" & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " ORDER BY NUM"
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NUM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_PAGE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RES_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MONTH_CNT"))
'            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_TOT"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT"), ggAmtOfMoney.DecPoint, 0)  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CODE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDI_CODE"))
            lgstrData = lgstrData & Chr(11) & ""
            
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
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)  
End Sub	
 
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubAutoQuery()
    Dim iDx,iSelCount
    Dim strWhere
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    strWhere = "SELECT " & FilterVar(lgKeyStream(0), "''", "S") & " comp_cd, a.emp_no  emp_no, a.name  name, " & _
				"	a.res_no + '0' emp_res_no,	ISNULL(c.work_month, 3) work_month, " & _
				"	d.income_tot_amt + d.non_tax1 + d.non_tax5 - isnull(t.before_tot_amt,0) -isnull(minus_tot  ,0)  tot_amt, " & _
					FilterVar(lgKeyStream(4), "''", "S") & "  jisa_code, " & FilterVar(lgKeyStream(5), "" & FilterVar("03", "''", "S") & "", "S") & " com_cd " & _
				" FROM hdf020t b left outer join haa010t a on b.emp_no = a.emp_no " &_
				"	 left outer join (select emp_no,CASE  WHEN count(distinct wk_yymm) < 3 THEN 3  ELSE count(distinct wk_yymm)  END AS work_month "&_
				"				 from hca090t"&_
				"				where wk_yymm LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " and wk_day >= 20"&_
				"				group by emp_no"&_
				"				) as c on b.emp_no = c.emp_no "&_
				"	 left outer join (select hca090t.emp_no, sum(isnull(tax_amt,0)+isnull(non_tax1,0)+isnull(non_tax5,0)) minus_tot"&_
				"				 from hca090t left outer join hdf070t on hca090t.emp_no =hdf070t.emp_no and hca090t.wk_yymm =hdf070t.pay_yymm"&_
				"				where wk_yymm LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " and wk_day < 20"&_
				"					and hdf070t.prov_type not in ('P','Q') " &_				
				"				group by hca090t.emp_no"&_
				"				) as f on b.emp_no = f.emp_no "&_				
				"	 left outer join (select emp_no, sum(isnull(a_pay_tot_amt,0) + isnull(a_bonus_tot_amt,0) + isnull(a_after_bonus_amt,0)) as before_tot_amt "&_
				"				from hfa040t"&_
				"				 where year_yy = " & FilterVar(lgKeyStream(1), "''", "S") &_
				"				group by emp_no) as t on b.emp_no = t.emp_no "&_
				"	left outer join hfa050t d on  b.emp_no = d.emp_no"&_
               " WHERE  a.sect_cd LIKE " & FilterVar(lgKeyStream(3), "'%'", "S") & _
               "  AND d.year_yy = " & FilterVar(lgKeyStream(1), "''", "S") & _
               "  AND ISNULL(b.anut_acq_dt, a.entr_dt) <= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & _
               "  AND a.entr_dt <= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") &_
               "  AND b.anut_grade IS NOT NULL AND b.anut_loss_dt IS NULL ORDER BY b.anut_acq_dt, b.anut_no"
 
 'Response.Write strWhere
 'Response.End
 
    If 	FncOpenRs("R",lgObjConn,lgObjRs,strWhere,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        lgstrData = ""
        iDx = 1
      
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UCase(ConvSPChars(lgObjRs("comp_cd")))
            lgstrData = lgstrData & Chr(11) & (String((5-Len(Cstr(lgStrPrevKey+iDx))),"0") & (lgStrPrevKey+iDx))
            lgstrData = lgstrData & Chr(11) & "000"
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_res_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("work_month"))
'			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tot_amt"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("tot_amt"), ggAmtOfMoney.DecPoint, 0)  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("jisa_code")   )         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("com_cd"))
            lgstrData = lgstrData & Chr(11) & "000"

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

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

    arrColVal = Split(arrRowVal(0), gColSep)   
    
    If  arrColVal(0) = "C" Then
		Call SubBizSaveMultiDelete(arrColVal)        
	End If

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
'            Case "D"
'                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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

    lgStrSQL = "INSERT INTO HDB040T(YEAR_YY, BIZ_AREA_NUM, NUM, BIZ_PAGE, RES_NO, NAME, MONTH_CNT, PAY_TOT, SUB_CODE, EDI_CODE, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
'    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)					& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")     & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDB040T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MONTH_CNT			= "     & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
'    lgStrSQL = lgStrSQL & " PAY_TOT			= "     & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " PAY_TOT				= "     & UNIConvNum(arrColVal(5),0)					& ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO			= "     & FilterVar(gUsrId, "''", "S")					& ","
    lgStrSQL = lgStrSQL & " ISRT_DT				= "     & FilterVar(lgGetSvrDateTime, "''", "S")
    
    lgStrSQL = lgStrSQL & " WHERE YEAR_YY		= " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND RES_NO			= " & FilterVar(UCase(arrColVal(3)), "''", "S")
'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = "DELETE  HDB040T"
    lgStrSQL = lgStrSQL & " WHERE YEAR_YY			= " & FilterVar(UCase(arrColVal(2)), "''", "S")
'Response.Write lgStrSQL
'Response.End
	lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizFile
' Desc : Save Data 
'============================================================================================================
Sub SubBizFile()
    Dim iDx
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = "SELECT BIZ_AREA_NUM, NUM, BIZ_PAGE, RES_NO, NAME, MONTH_CNT, PAY_TOT, SUB_CODE, EDI_CODE"
	lgStrSQL = lgStrSQL & " FROM HDB040T "
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY =" & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " ORDER BY NUM"

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(lgObjRs("BIZ_AREA_NUM")		,"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("NUM")		,"","0",5,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("BIZ_PAGE")		,"","0",3,"RIGHT")
            lgstrData = lgstrData & ConvSPChars(lgObjRs("RES_NO"))
            lgstrData = lgstrData & SetFixSrting((String((2-Len(Cstr(lgObjRs("MONTH_CNT")))),"0") & lgObjRs("MONTH_CNT"))		,"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting((String((9-Len(Cstr(lgObjRs("PAY_TOT")))),"0") & lgObjRs("PAY_TOT"))  		,"","0",9,"RIGHT")    
            lgstrData = lgstrData & ConvSPChars(lgObjRs("SUB_CODE")  )          
            lgstrData = lgstrData & ConvSPChars(lgObjRs("EDI_CODE"))
            lgstrData = lgstrData & "000"
            lgstrData = lgstrData & Chr(13) & Chr(10)

		    lgObjRs.MoveNext
            iDx =  iDx + 1

              
        Loop 
'Response.Write "lgstrData:"  & lgstrData        
' Response.End         
    End If
    
    lgStrPrevKey = ""

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

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
' Name : SetFixSrting(입력값,비교문자,대체문자,고정길이,문자정렬방향)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '입력값이 존재하지않을경우 입력값의 길이를 0으로 한다.
        Cnt = 0     
    Else                                  '입력값이 존재하면서 한글일경우 
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2                                                  '한글부분만 길이를 각각 2로한다.
            Else
                MCnt = MCnt + 1    
            End If
        Next
        Cnt = MCnt
                         
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '비교문자가 없을경우 
        Else                                                                         '비교문자가 존재할경우 비교문자를 뺀 나머지를 입력값으로한다.
            ExSymbol = Split(InValue,ComSymbol)
            If UBound(ExSymbol) > 0 Then
                iDx = UBound(ExSymbol)
                For i = 0 To iDx
                    strSplit = strSplit & ExSymbol(i)
                Next
                InValue = strSplit
            End If               
        End If        
    End If        
    
    If InPos = "" Then                                                              '고정길이가 정해지지 않을 경우 입력문자 길이가 고정길이가 된다.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '왼쪽정렬일경우(defalut) 고정길이 보다 작은 길이의 문자가 입력되면 나머지 공백(default)부분을 대체문자로 체운다.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '오른쪽정렬 
        If InPos > Cnt Then                                                           ' ex:     hi
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = 1 To (InPos - Cnt)
                strTemp = strTemp & strFix         
            Next
            InValue = strTemp & InValue
        End If
    End If
    SetFixSrting = InValue
End Function

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
       Case "7"                                                         '☜ : Delete
		  Dim SF
		  On Error Resume Next
		  Set SF = CreateObject("uni2kCM.SaveFile")
		  Call SF.SaveTextFile("<%= strFilePath %>")

		  Set SF = Nothing
       Case "5"
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
