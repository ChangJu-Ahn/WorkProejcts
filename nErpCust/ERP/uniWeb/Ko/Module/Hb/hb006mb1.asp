<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey,lgStrPrevKey1,lgStrPrevKey2    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")
    
    Dim strEmpno
    Dim strNo
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    If lgCurrentSpd = "A" Then    
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
	elseif lgCurrentSpd = "B" Then    
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)
	elseif lgCurrentSpd = "C" Then   
		lgStrPrevKey2 = UNICInt(Trim(Request("lgStrPrevKey2")),0)	
	end if	 	
        
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()        
	Dim strWhere    
	Dim pComp

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            strWhere = FilterVar(lgKeyStream(6), "''", "S")
            
        Case "B"

            strWhere = " HAA011T.PROV_TYPE = 'Y' "
			strWhere = strWhere & " AND HAA011T.YEAR_AREA_CD  like " & FilterVar(lgKeyStream(7), "''", "S")
			strWhere = strWhere & " AND pay_yymm >=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) "
			strWhere = strWhere & " AND pay_yymm <=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3*" & lgKeyStream(3) & "),2) "
			            
        Case "C"
            strWhere = " HAA011T.PROV_TYPE = 'Y' "
			strWhere = strWhere & " AND HAA011T.YEAR_AREA_CD  like " & FilterVar(lgKeyStream(7), "''", "S")
			strWhere = strWhere & " AND pay_yymm >=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) "
			strWhere = strWhere & " AND pay_yymm <=" & FilterVar(lgKeyStream(4), "''", "S") 
			strWhere = strWhere & "  + right( '0' +convert(varchar(2),3*" & lgKeyStream(3) & "),2) "
        
    End Select
    
    Call SubMakeSQLStatements("MR",strWhere,"x","")                              '☆: Make sql statements    
    Call SubBizQueryMulti()    
End Sub	
 
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    
    Dim iDx
    Dim DFnm        
    Dim li_biz_own_rgst_no
    Dim Oldres_no,Cwork_no
    Dim i,strDNO
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""    
        lgStrPrevKey1 = ""    
        lgStrPrevKey2 = ""      
'        Call SetErrorStatus()
    Else    
       Select Case UCase(Trim(lgCurrentSpd))
             Case "A"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)    
    		 Case "B"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)    
    		 Case "C"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey2)    
        End Select    		    		 
     
        lgstrData = ""
        Oldres_no = ""
        Cwork_no = 0
        li_biz_own_rgst_no = Trim(lgKeyStream(4))        
        iDx = 1
        
        Do While Not lgObjRs.EOF
            Select Case UCase(Trim(lgCurrentSpd))
                 Case "A"
                     If Trim(li_biz_own_rgst_no) = "" Or Trim(li_biz_own_rgst_no) <> Trim(replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-","")) Then 
                         li_biz_own_rgst_no = Trim(replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-",""))
                         li_biz_own_rgst_no = Left(li_biz_own_rgst_no,7) & "." & Right(li_biz_own_rgst_no,3)
                     End If

                     Call CommonQueryRs("count(*) ","HFA100T","year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                                          
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dcl_date"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("p_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mag_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hometax_id"))  
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("taxpgm_cd"))                       
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_DEPT_NM"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_NAME"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_TEL"))
                     lgstrData = lgstrData & Chr(11) & Replace(lgF0, Chr(11), "")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("term_code"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(" ")
                Case "B"   
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & iDx
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("repre_nm"))
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("com_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("com_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_com_no"))
                     
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("base_year"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("term_code"))
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("emp_cnt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("prov_tot_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("income_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("res_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(" ")
                Case "C"
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & iDx
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-","")
                     
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NM"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("for_type")) 
                     
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("retire_month"), 0,0)        
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("prov_tot_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("income_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("res_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(" ")
           End Select
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1      
            If iDx > C_SHEETMAXROWS_D Then
				Select Case UCase(Trim(lgCurrentSpd))
				    Case "A"
						lgStrPrevKey = lgStrPrevKey + 1
				    Case "B"
						lgStrPrevKey1 = lgStrPrevKey1 + 1
				    Case "C"
						lgStrPrevKey2 = lgStrPrevKey2 + 1
				End Select					
               Exit Do
            End If                       
                    
        Loop         
        If Trim(lgCurrentSpd) = "A" then
            DFnm = "C:\I" & li_biz_own_rgst_no       
%>
<SCRIPT LANGUAGE=VBSCRIPT>
    parent.frm1.txtFile.value = "<%=DFnm%>"
</SCRIPT>
<%      End If
    End If   	
    If iDx <= C_SHEETMAXROWS_D Then
		Select Case UCase(Trim(lgCurrentSpd))
		    Case "A"
		       lgStrPrevKey = ""
		    Case "B"
		       lgStrPrevKey1 = ""
		    Case "C"
		       lgStrPrevKey2 = ""
		End Select		       
    End If           
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub    
 
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))
                Case "A"
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
					
                    lgStrSQL = " SELECT top " & iSelCount & " 'A' record_type,'28' data_type,"		'/* 레코드구분(A로고정), 자료구분:28으로 고정 */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S") & " dcl_date,"		'/* 제출연월일 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & " p_type,"			'/* 제출자(대리인)구분 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S") & "  mag_no,"				'/* 세무대리인관리번호 */
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  "				'/* 2004 hometax id */  
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  "								'/* 2004 세무프로그램코드 기타 */                                                             
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_rgst_no,"							'/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm) biz_area_nm,"	'/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM, WORKER_NAME, WORKER_TEL,"			'담당자 부서/담당자명/담당자 전화번호 2004   
					lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & " term_code"
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd = " & pCode

'Response.Write lgStrSQL
'Response.End 

                Case "B" 
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
					                
                    lgStrSQL = " SELECT top " & iSelCount & " hfa100t.year_area_cd  singo_org_cd,"	
                    lgStrSQL = lgStrSQL & " 'B' record_type,'28' data_type, "							'/* 레코드구분,자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"									'/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"							'/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"							'/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"								'/* 대표자(성명) */
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  com_rgst_no,"						'/* 주민(법인)등록번호 */
					
					lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(4), "''", "S") & " base_year, "			'/* 귀속연도 */
					lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & " term_code, "			'/* 귀속분기 */
					
					lgStrSQL = lgStrSQL & " count(distinct hdf071t.emp_no)		emp_cnt,"				'/* 일용근로인원수 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.prov_tot_amt))	prov_tot_amt,"			'/* 총지급액 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.income_tax))		income_tax,"			'/* 소득세 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.res_tax))			res_tax"				'/* 주민세 */
                    lgStrSQL = lgStrSQL & " FROM hdf071t left outer join haa011t on hdf071t.emp_no = haa011t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hdf020t on hdf071t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hfa100t on haa011t.year_area_cd = hfa100t.year_area_cd"
					lgStrSQL = lgStrSQL & " WHERE " & pCode 

					lgStrSQL = lgStrSQL & " GROUP BY hfa100t.year_area_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no,"
					lgStrSQL = lgStrSQL & " hfa100t.year_area_nm,"
					lgStrSQL = lgStrSQL & " hfa100t.repre_nm,"
					lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no"
'Response.Write lgStrSQL
'Response.End
                Case "C" 
						
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey2 + 1
					                
                    lgStrSQL = " SELECT top " & iSelCount & " hfa100t.year_area_cd  singo_area_cd," 
                    lgStrSQL = lgStrSQL & " 'C' record_type, '28' data_type, "					'/* 레코드구분/자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"							'/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"					'/* 사업자등록번호 */
                    
                    lgStrSQL = lgStrSQL & " haa011t.res_no  res_no,"							'/* 주민(법인)등록번호 */
                    lgStrSQL = lgStrSQL & " haa011t.EMP_NM,"										'/* 성명 */  
                    lgStrSQL = lgStrSQL & " CASE WHEN haa011t.NATIVE_CD = 'KR' THEN '1' ELSE '9' END for_type,"	'/* 내외국인구분코드 */
                    
                    lgStrSQL = lgStrSQL & " case when haa011t.RETIRE_DT >=" & FilterVar(lgKeyStream(4), "''", "S") 
                    lgStrSQL = lgStrSQL & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & " -2),2) +'01' "
                    lgStrSQL = lgStrSQL & " AND haa011t.RETIRE_DT < dateadd(month,1," & FilterVar(lgKeyStream(4), "''", "S") 
                    lgStrSQL = lgStrSQL & "  + right( '0' +convert(varchar(2),3* " & lgKeyStream(3) & "),2) +'01' ) THEN month(haa011t.RETIRE_DT)"
                    lgStrSQL = lgStrSQL & " ELSE 3*" & lgKeyStream(3) & " END retire_month,"			'/* 근로종료월 */                    

                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.prov_tot_amt))	prov_tot_amt,"			'/* 총지급액 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.income_tax))		income_tax,"			'/* 소득세 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hdf071t.res_tax))			res_tax "				'/* 주민세 */

                    lgStrSQL = lgStrSQL & " FROM hdf071t left outer join haa011t on hdf071t.emp_no = haa011t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hdf020t on hdf071t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & "				 left outer join hfa100t on haa011t.year_area_cd = hfa100t.year_area_cd"                         
                    lgStrSQL = lgStrSQL & " WHERE " & pCode 
                    
					lgStrSQL = lgStrSQL & " GROUP BY hfa100t.year_area_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd,"
					lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no,"
					lgStrSQL = lgStrSQL & " haa011t.res_no,haa011t.EMP_NM, haa011t.NATIVE_CD,haa011t.RETIRE_DT"
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
 
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                Select Case Trim("<%=lgCurrentSpd%>")
                    Case "A"
                        .ggoSpread.Source     = .frm1.vspdData
                        .lgStrPrevKey    = "<%=lgStrPrevKey%>"
						if .topleftOK then
							.DBQueryOk
						else
							.lgCurrentSpd = "B"						
							.DBQuery
						end if
                        
                    Case "B"
                        .ggoSpread.Source     = .frm1.vspdData1
                        .lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
						if .topleftOK then
							.DBQueryOk
						else
							.lgCurrentSpd = "C"						
							.DBQuery
						end if
                        
                    Case "C"
                        .ggoSpread.Source     = .frm1.vspdData2
                        .lgStrPrevKey2    = "<%=lgStrPrevKey2%>"
		                .DBQueryOk                          
                End Select
               .ggoSpread.SSShowData "<%=lgstrData%>"
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
