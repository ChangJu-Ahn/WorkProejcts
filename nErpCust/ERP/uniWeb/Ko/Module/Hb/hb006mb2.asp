<%@ LANGUAGE=VBSCript%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncServer.asp" -->    
<!-- #Include file="../../inc/lgsvrvariables.inc" -->   
<!-- #Include file="../../inc/incServeradodb.asp" -->   
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/uni2kcm.inc" -->
<%
  On Error Resume Next	
  Err.Clear

    Dim AlgObjRs,BlgObjRs,ClgObjRs
    Dim BiDx
    Dim strFilePath,strMode
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm    
    Call HideStatusWnd                                                              '☜: Hide Processing message
    BiDx = 1

    lgErrorStatus     = "NO"													    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '☜: "M"(Spread #1) "S"(Spread #2)
    
    strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
    lgstrData = ""
 
    Call SubOpenDB(lgObjConn)      
            
    Select Case strMode
	    Case CStr(UID_M0001)   	                                                          'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")  
                Fnm = Fso.GetFileName(lgKeyStream(0))                
                FPnm = Server.MapPath("../../files/u2000/" & Fnm)       
                
                Call SubBizQuery()

            If UCase(Trim(lgErrorStatus)) <> "YES" Then
                Set CTFnm = Fso.CreateTextFile (FPnm,True)                              'text를 저장할 File을 생성            
                
                CTFnm.Write lgstrData                                                   'Text 내용부분                       
                DFnm = Fso.GetFileName(FPnm)            
                CTFnm.close    
                Set CTFnm = nothing
            Else
                Call DisplayMsgBox("700100", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                Call SetErrorStatus() 
            End If
            Set Fso = nothing           
            
            Call HideStatusWnd           
            
%>
    <SCRIPT LANGUAGE=VBSCRIPT>
				parent.subVatDiskOK("<%=DFnm%>")
	</SCRIPT>
<%
    Case CStr(UID_M0002)

	    Err.Clear 

	    Call HideStatusWnd

	    strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" _
	    			   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
	    strFilePath = strFilePath  & "files/u2000/"
	    strFilePath = strFilePath & Request("txtFileName")

End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()        
	Dim strWhere    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                           '☜: Clear Error status
	
    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            strWhere = FilterVar(lgKeyStream(6), "''", "S")
            strWhere = strWhere & CQuery
            
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements                   
            Call SubBizQueryMulti()    
            
        Case "B"        
            strWhere = " HAA011T.PROV_TYPE = 'Y' "
			strWhere = strWhere & " AND HAA011T.YEAR_AREA_CD  like " & FilterVar(lgKeyStream(7), "''", "S")
			strWhere = strWhere & " AND pay_yymm >=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) "
			strWhere = strWhere & " AND pay_yymm <=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3*" & lgKeyStream(3) & "),2) "
            strWhere = strWhere & CQuery
 
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements                   
            
        Case "C"
            strWhere = " HAA011T.PROV_TYPE = 'Y' "
			strWhere = strWhere & " AND HAA011T.YEAR_AREA_CD  like " & FilterVar(lgKeyStream(7), "''", "S")
			strWhere = strWhere & " AND pay_yymm >=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) "
			strWhere = strWhere & " AND pay_yymm <=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3*" & lgKeyStream(3) & "),2) "             
            strWhere = strWhere & " AND hfa100t.own_rgst_no =  " & FilterVar(BlgObjRs("biz_rgst_no"), "''", "S")'by 20060208

            Call SubMakeSQLStatements("MR",strWhere,"x","")                              '☆: Make sql statements                           
    End Select       
End Sub	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    Call ASubBizQueryMulti()        
End Sub    
'============================================================================================================
' Name : ASubBizQueryMulti()
' Desc : Query ASheet Data from Db
'============================================================================================================
Sub ASubBizQueryMulti()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                             '☜: Clear Error status

    lgstrData = ""
    If FncOpenRs("R",lgObjConn,AlgObjRs,lgStrSQL,"X","X") = False Then
       Call SetErrorStatus("")
    Else        
        Do While Not AlgObjRs.EOF        
            Call CommonQueryRs("count(*) ","HFA100T","year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

            lgstrData = lgstrData & SetFixSrting(AlgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("dcl_date"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("p_type"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("mag_no"),"","",6,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("hometax_id"),"","",20,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("taxpgm_cd"),"","",4,"")
            lgstrData = lgstrData & SetFixSrting(replace(AlgObjRs("biz_rgst_no"),"-",""),"-","",10,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("biz_area_nm"),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_DEPT_NM"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_NAME"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("WORKER_TEL"),"","",15,"")
            lgstrData = lgstrData & SetFixSrting(Replace(lgF0, Chr(11), ""),"","0",5,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("term_code"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting("","","",4," ")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            
            If Cdbl(ConvSPChars(AlgObjRs("b_count"))) > 0 Then                
                lgCurrentSpd = "B"
                Call BSubBizQueryMulti()
                lgCurrentSpd = "A"
            End If
            AlgObjRs.MoveNext
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,AlgObjRs,Err)
    Call SubCloseRs(AlgObjRs)    
End Sub
'============================================================================================================
' Name : BSubBizQueryMulti()
' Desc : Query BSheet Data from Db
'============================================================================================================
Sub BSubBizQueryMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
                                                                            '☜: Clear Error status    

    Call SubBizQuery()
    If 	FncOpenRs("R",lgObjConn,BlgObjRs,lgStrSQL,"X","X") = False Then
        Call SetErrorStatus()
    Else    
        Do While Not BlgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(BiDx,"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(BlgObjRs("biz_rgst_no"),"-",""),"-","",10,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("biz_area_nm"),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("repre_nm"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(replace(BlgObjRs("com_rgst_no"),"-",""),"-","",13,"")
            
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("base_year"),"","",4,"")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("term_code"),"","",1,"")
            
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("emp_cnt"),"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("prov_tot_amt"),"","0",15,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("income_tax"),"","0",15,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("res_tax"),"","0",15,"RIGHT")
            lgstrData = lgstrData & SetFixSrting("","","",19," ")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            IF Cdbl(ConvSPChars(BlgObjRs("com_no"))) > 0 Then                
                lgCurrentSpd = "C"
                Call CSubBizQueryMulti()
                lgCurrentSpd = "B"
            End If
            BiDx =  BiDx + 1
            BlgObjRs.MoveNext
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,BlgObjRs,Err)
    Call SubCloseRs(BlgObjRs)
End Sub
'============================================================================================================
' Name : CSubBizQueryMulti()
' Desc : Query CSheet Data from Db
'============================================================================================================
Sub CSubBizQueryMulti()
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
                                                                             '☜: Clear Error status
    Call SubBizQuery()
    If 	FncOpenRs("R",lgObjConn,ClgObjRs,lgStrSQL,"X","X") = False Then
        Call SetErrorStatus()
    Else    
        iDx = 1
        Do While Not ClgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("record_type"),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("data_type"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("tax"),"","",3,"")
            lgstrData = lgstrData & SetFixSrting(iDx,"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(ClgObjRs("biz_rgst_no"),"-",""),"-","",10,"")

            lgstrData = lgstrData & SetFixSrting(replace(ClgObjRs("res_no"),"-",""),"-","",13,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("EMP_NM"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("for_type"),"","0",1,"RIGHT")

            lgstrData = lgstrData & SetFixSrting(ClgObjRs("retire_month"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("prov_tot_amt"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("income_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("res_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting("","","",73," ")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            ClgObjRs.MoveNext
            iDx = iDx + 1
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,ClgObjRs,Err)
    Call SubCloseRs(ClgObjRs)
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
                                                                         '☜: Clear Error status
    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))                
                Case "A"
                    lgStrSQL = " SELECT 'A' record_type,'28' data_type,"		'/* 레코드구분(A로고정), 자료구분:28으로 고정 */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"					'/* 세무서 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S") & " dcl_date,"		'/* 제출연월일 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & " p_type,"			'/* 제출자(대리인)구분 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S") & "  mag_no,"		'/* 세무대리인관리번호 */  
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  "					'/* 2004 hometax id */                                       
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  "									'/* 2004 세무프로그램코드 기타 */                      
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_rgst_no,"								'/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm)   biz_area_nm,"		'/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM, WORKER_NAME, WORKER_TEL,"				'담당자 부서/담당자명/담당자 전화번호 2004   
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & " term_code"
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd = " & pCode

                Case "B" 
                    lgStrSQL = " SELECT hfa100t.year_area_cd  singo_org_cd,"
                    lgStrSQL = lgStrSQL & " 'B' record_type,'28' data_type, "					'/* 레코드구분,자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"							'/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"					'/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"					'/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"								'/* 대표자(성명) */                    
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  com_rgst_no,"               '/* 주민(법인)등록번호 */

					lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(4), "''", "S") & " base_year, "			'/* 귀속연도 */
					lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & " term_code, "			'/* 귀속분기 */
					
					lgStrSQL = lgStrSQL & " count(distinct hdf071t.emp_no)		emp_cnt,"				'/* 일용근로인원수 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.prov_tot_amt))	prov_tot_amt,"			'/* 총지급액 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.income_tax))		income_tax,"			'/* 소득세 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.res_tax))			res_tax"				'/* 주민세 */
                    lgStrSQL = lgStrSQL & " FROM hdf071t left outer join HAA011T on hdf071t.emp_no = HAA011T.emp_no"
                    'lgStrSQL = lgStrSQL & "				 left outer join hdf020t on hdf071t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hfa100t on HAA011T.year_area_cd = hfa100t.year_area_cd"
					lgStrSQL = lgStrSQL & " WHERE " & pCode 

					lgStrSQL = lgStrSQL & " GROUP BY hfa100t.year_area_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no,"
					lgStrSQL = lgStrSQL & " hfa100t.year_area_nm,"
					lgStrSQL = lgStrSQL & " hfa100t.repre_nm,"
					lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no"                    
                
                Case "C" 

                    lgStrSQL = " SELECT hfa100t.year_area_cd  singo_area_cd," 
                    lgStrSQL = lgStrSQL & " 'C' record_type, '28' data_type, "					'/* 레코드구분/자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"							'/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"					'/* 사업자등록번호 */

                    lgStrSQL = lgStrSQL & " HAA011T.res_no  res_no,"							'/* 주민(법인)등록번호 */
                    lgStrSQL = lgStrSQL & " HAA011T.EMP_NM,"										'/* 성명 */  
                    lgStrSQL = lgStrSQL & " CASE WHEN HAA011T.NATIVE_CD = '1' THEN '1' ELSE '9' END for_type,"	'/* 내외국인구분코드 */
                    
                    lgStrSQL = lgStrSQL & " case when HAA011T.RETIRE_DT >=" & FilterVar(lgKeyStream(4), "''", "S") 
                    lgStrSQL = lgStrSQL & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) +'01' "
                    lgStrSQL = lgStrSQL & " AND HAA011T.RETIRE_DT < dateadd(month,1," & FilterVar(lgKeyStream(4), "''", "S") 
                    lgStrSQL = lgStrSQL & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & "),2) +'01' ) THEN month(HAA011T.RETIRE_DT)"
                    lgStrSQL = lgStrSQL & " ELSE 3*" & lgKeyStream(3) & " END retire_month,"			'/* 근로종료월 */  
                    
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.prov_tot_amt))	prov_tot_amt,"			'/* 총지급액 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.income_tax))		income_tax,"			'/* 소득세 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.res_tax))			res_tax "				'/* 주민세 */

                    lgStrSQL = lgStrSQL & " FROM hdf071t left outer join HAA011T on hdf071t.emp_no = HAA011T.emp_no"
                    'lgStrSQL = lgStrSQL & "				 left outer join hdf020t on hdf071t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hfa100t on HAA011T.year_area_cd = hfa100t.year_area_cd"                         
                    lgStrSQL = lgStrSQL & " WHERE " & pCode 
                    
					lgStrSQL = lgStrSQL & " GROUP BY hfa100t.year_area_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no,"
					lgStrSQL = lgStrSQL & " HAA011T.res_no,HAA011T.EMP_NM, HAA011T.NATIVE_CD,HAA011T.RETIRE_DT"

' Response.Write lgStrSQL
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear 
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
    Else

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

%>

<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>
