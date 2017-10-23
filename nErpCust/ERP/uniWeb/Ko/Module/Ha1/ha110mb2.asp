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
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")  
                Fnm = Fso.GetFileName(lgKeyStream(4))                
                FPnm = Server.MapPath("../../files/u2000/" & Fnm)       
                
                Call SubBizQuery("")

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

	    strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
	    			   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
'	    strFilePath = strFilePath  & "files/" & gCompany & "/"
        strFilePath = strFilePath  & "files/u2000/"  '2002.02.01 /files 에는 현재 u2000만 존재:나중에 공통쪽 변경되면 수정해야 함.
	    strFilePath = strFilePath & Request("txtFileName")

End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(CQuery)        
Dim strWhere    
Dim pComp
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            pComp = "="
            strWhere = FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & CQuery
            
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements       
            Call SubBizQueryMulti()    
        Case "B"
            pComp = ""
            strWhere = FilterVar(lgKeyStream(5),"NULL", "S") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery
            strWhere = strWhere & " GROUP BY hfa100t.year_area_cd,"                                                       '바뀐 Query
            strWhere = strWhere & " hfa100t.tax_biz_cd,"
            strWhere = strWhere & " hfa100t.own_rgst_no,"
            strWhere = strWhere & " hfa100t.year_area_nm,"
            strWhere = strWhere & " hfa100t.repre_nm,"
            strWhere = strWhere & " hfa100t.co_own_rgst_no"

            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements                   
        Case "C"        
            pComp = ""
            strWhere = replace(FilterVar(lgKeyStream(5),"NULL", "S"),gComDateType,"") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND haa010t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery

            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements                           
    End Select       
End Sub	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call ASubBizQueryMulti()        
End Sub    
'============================================================================================================
' Name : ASubBizQueryMulti()
' Desc : Query ASheet Data from Db
'============================================================================================================
Sub ASubBizQueryMulti()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
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
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("kr_code"),"","0",3,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("term_code"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(AlgObjRs("empty"),"","",1,"")
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
    Err.Clear                                                                        '☜: Clear Error status    
    
    Call SubBizQuery("")
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
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("com_no"),"","0",7,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("old_com_no"),"","0",7,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tot_prov_amt"),"","0",14,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("deci_income_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("tot_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("deci_res_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("deci_farm_tax"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("deci_sum"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(BlgObjRs("empty"),"","",132,"")
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
    Err.Clear                                                                        '☜: Clear Error status
    
    Call SubBizQuery("")
    
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
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_com_no"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hdf020t_res_flag"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_nat_cd"),"","",2,"")   '2002 거주지국코드 
            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("haa010t_entr_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("hga070t_retire_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_name"),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("for_type"),"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(ClgObjRs("res_no"),"-",""),"-","",13,"")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("retire_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_honor_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_corp_insur"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_tot_prov_amt"),"","0",10,"RIGHT")
            
            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("entr_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("retire_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_tot_duty_mm"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_entr_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_retire_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_duty"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("d_duty"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_duty_cnt"),"","0",2,"RIGHT")
            
'명예퇴직 중간정산자 2005-06-30

            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("h_entr_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(UNIConvDateToYYYYMMDD(ClgObjRs("h_retire_dt"),gServerDateFormat,""),gServerDateType,"0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("tot_duty_mm2"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_old_retire_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_retire_dt"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_old_duty"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_d_duty"),"","0",4,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_duty_cnt"),"","0",2,"RIGHT")
           
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("retire_tot_prov_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("income_sub"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_avr_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_avr_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("retire_sub"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("deci_tax"),"","0",10,"RIGHT")            

            lgstrData = lgstrData & SetFixSrting(ClgObjRs("honor_amt"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_income_sub"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_hga070t_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_hga070t_avr_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_hga070t_avr_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_hga070t_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_retire_sub"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("h_deci_tax"),"","0",10,"RIGHT")   

            lgstrData = lgstrData & SetFixSrting(ClgObjRs("tot_income"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_income_sub"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_hga070t_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_hga070t_avr_tax_std"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_hga070t_avr_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_hga070t_calc_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_retire_sub"),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("t_deci_tax"),"","0",10,"RIGHT")     
            
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_deci_income_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hga070t_deci_res_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("deci_farm_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("deci_sum"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_income_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_res_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("hfa050t_old_farm_tax"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("old_sum"),"","0",10,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("empty_2"),"","",1,"")    
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
	Dim re_tax_sub1,re_sub_limit, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))                
                Case "A"
                    lgStrSQL = " SELECT " & FilterVar("A", "''", "S") & "  record_type,"	'/* 레코드구분 : A로 고정 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"	'/* 자료구분 : 22으로 고정 */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(2), "''", "S") & " dcl_date,"	'/* 제출연월일 -> 입력변수 */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & " p_type,"		'/* 제출자(대리인)구분 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & "  mag_no,"	'/* 세무대리인관리번호 */  
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  "		'/* 2004 hometax id */                                       
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  "						'/* 2004 세무프로그램코드 기타 */                      
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_rgst_no,"			'/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " year_area_nm  biz_area_nm,"			'/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM,  "		'담당자 부서 2004   
                    lgStrSQL = lgStrSQL & " WORKER_NAME,  "			'담당자명 2004   
                    lgStrSQL = lgStrSQL & " WORKER_TEL,  "			'담당자 전화번호 2004  

                    lgStrSQL = lgStrSQL & " " & FilterVar("101", "''", "S") & " kr_code,"                                 '/* 사용한글코드 : 101로 고정 */
'                    lgStrSQL = lgStrSQL & " " & FilterVar(lgKeyStream(1), "''", "S") & " term_code,"                                 '/* 제출대상기간코드 -> 입력변수 */
                    lgStrSQL = lgStrSQL & " '1'  term_code,"
                    
                    lgStrSQL = lgStrSQL & " SPACE(371) empty"                               '/* 공란 */ 
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd" & pComp & " " &  pCode
'Response.Write "lgStrSQL2:" & lgStrSQL  

                Case "B" 
                    lgStrSQL = " SELECT hfa100t.year_area_cd  singo_org_cd,"
                    lgStrSQL = lgStrSQL & " " & FilterVar("B", "''", "S") & "  record_type,"                                   '/* 레코드구분 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"                                    '/* 자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"               '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"               '/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  com_rgst_no,"                '/* 주민(법인)등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"                     '/* 대표자(성명) */
                    lgStrSQL = lgStrSQL & " COUNT(hga070t.emp_no) com_no,"                      '/* 주(현)제출건수(C레코드수) */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0", "''", "S") & "  old_com_no,"                                    '/* 종(전)제출건수(D레코드수) -> 0으로 무조건 고정 */ 
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.tot_prov_amt))  tot_prov_amt,"    '/* 소득금액 총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2))  deci_income_tax,"  '/* 소득세결정세액총계 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000000", "''", "S") & " tot_tax,"                               '/* 법인세결정세액총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_res_tax)) deci_res_tax,"         '/* 주민세결정세액총계 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000000", "''", "S") & " deci_farm_tax,"                         '/* 농특세결정세액총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2) + FLOOR(hga070t.deci_res_tax)) deci_sum,"  '/* 결정세액총계 */
                    lgStrSQL = lgStrSQL & " SPACE(1) d_code,"                                   '/* 자료수정코드 */
                    lgStrSQL = lgStrSQL & " SPACE(352) empty"                                   '/* 공란 */
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hga070t,"
                    lgStrSQL = lgStrSQL & " hfa100t"                                              '바뀐 Query
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hga070t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_yy *= hfa050t.year_yy"
                    lgStrSQL = lgStrSQL & " AND hga070t.emp_no *= hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
'                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hga070t.honor_retire_flag=" & FilterVar("Y", "''", "S")     
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_dt BETWEEN" & pComp & " " &  pCode
'Response.EndResponse.Write "lgStrSQL2:" & lgStrSQL                      

                Case "C" 
                    call  CommonQueryRs(" RE_TAX_SUB1 , RE_SUB_LIMIT "," hda000t "," 1=1" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					re_tax_sub1 = Replace(Trim(lgF0), Chr(11), "")
					re_tax_sub1 = CInt(re_tax_sub1) / 100
					re_sub_limit = Replace(Trim(lgF1), Chr(11), "")
						
						
                    If re_tax_sub1 = "" OR re_tax_sub1 = "X" Then
						re_tax_sub1 = "0.25"
						re_sub_limit = "120000"
					end if
                
                    lgStrSQL = " SELECT hfa100t.year_area_cd  singo_area_cd," 
                    lgStrSQL = lgStrSQL & " " & FilterVar("C", "''", "S") & "  record_type,"                                           '/* 레코드구분 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"                                            '/* 자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"                       '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00", "''", "S") & " old_com_no,"                                           '/* 종(전)근무처수 -> 무조건 '00'으로 고정 */
                    lgStrSQL = lgStrSQL & " CASE WHEN (hdf020t.res_flag IS NULL OR hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " ) THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("2", "''", "S") & " END  hdf020t_res_flag,"  '/* 거주자구분코드 */
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd = " & FilterVar("KR", "''", "S") & ") THEN " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE haa010t.nat_cd END  haa010t_nat_cd,"  '2002 거주지국코드 
                    
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.entr_dt < convert(datetime," & FilterVar(lgKeyStream(8) & "-01-01", "''", "S")&") THEN " & FilterVar(lgKeyStream(8) & "-01-01" ,"NULL", "S")
                    lgStrSQL = lgStrSQL & " ELSE haa010t.entr_dt END haa010t_entr_dt,"  
                    lgStrSQL = lgStrSQL & " hga070t.RETIRE_DT  hga070t_retire_dt,"  

                    lgStrSQL = lgStrSQL & " haa010t.name  haa010t_name,"                                '/* 성명 */  
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.nat_cd = " & FilterVar("KR", "''", "S") & " THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("9", "''", "S") & " END for_type,"                                     '/* 내외국인구분코드 */
                    lgStrSQL = lgStrSQL & " haa010t.zip_cd zip,"                                        '/* 우편번호 : 2000년 연말정산 폐지 */
                    lgStrSQL = lgStrSQL & " haa010t.addr address,"                                      '/* 주소 : 2000년 연말정산 폐지 */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"                                    '/* 주민(법인)등록번호 */
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.retire_amt) + FLOOR(hga070t.etc_amt))  retire_amt,"  '/* 퇴직급여 */  
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.honor_amt)  hga070t_honor_amt,"                '/* 명예퇴직수당 또는 추가퇴직급여 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.corp_insur)  hga070t_corp_insur,"              '/* 단체퇴직보험금 */ 
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.retire_amt) + FLOOR(hga070t.etc_amt)) + FLOOR(hga070t.honor_amt)+FLOOR(hga070t.corp_insur)   hga070t_tot_prov_amt,"          '/* 계 */
                    lgStrSQL = lgStrSQL & " hga070t.entr_dt entr_dt,"                                    '/* 주(현)근무지입사연월일 */
                    lgStrSQL = lgStrSQL & " hga070t.retire_dt retire_dt,"                                '/* 주(현)근무지퇴사연월일 */
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetServeMonths(hga070t.entr_dt,hga070t.retire_dt) hga070t_tot_duty_mm," '/* 주(현)근무지근속월수  -2002.03.25*/
                    
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & " old_entr_dt,"                                     '======>/* 종(전)근무지입사연월일 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & " old_retire_dt,"                                   '======>/* 종(전)근무지퇴사연월일 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000", "''", "S") & " old_duty,"                                            '======>/* 종(전)근무지 근속월수 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000", "''", "S") & " d_duty,"                                              '======>/* 중복월수 */
                    
                    lgStrSQL = lgStrSQL & " ceiling(dbo.ufn_GetServeMonths(hga070t.entr_dt,hga070t.retire_dt) / 12.0) hga070t_duty_cnt,"  '/* 근속연수 -2002.03.25 */
                    

'명예퇴직 중간정산자 2005-06-30
                    lgStrSQL = lgStrSQL & " haa010t.entr_dt h_entr_dt,"                                    '/* 주(현)근무지입사연월일 */
                    lgStrSQL = lgStrSQL & " haa010t.retire_dt h_retire_dt,"                                '/* 주(현)근무지퇴사연월일 */
                    lgStrSQL = lgStrSQL & " hga070t.tot_duty_mm2," '/* 주(현)근무지근속월수  -2002.03.25*/

                    
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " h_old_entr_dt,"                                     '======>/* 종(전)근무지입사연월일 */
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " h_old_retire_dt,"                                   '======>/* 종(전)근무지퇴사연월일 */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " h_old_duty,"                                            '======>/* 종(전)근무지 근속월수 */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " h_d_duty,"                                              '======>/* 중복월수 */
                    
                    lgStrSQL = lgStrSQL & " ceiling(dbo.ufn_GetServeMonths(haa010t.entr_dt,haa010t.retire_dt) / 12.0) haa010t_duty_cnt,"  '/* 근속연수 -2002.03.25 */
'--------------                    
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.retire_amt) retire_tot_prov_amt,"            '/* 퇴직급여액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub) + FLOOR(hga070t.special_sub) income_sub,"  '/* 퇴직소득공제 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std)  hga070t_tax_std,"                     '/* 퇴직소득과표 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std)  hga070t_avr_tax_std,"             '/* 연평균과세표준 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax)  hga070t_avr_calc_tax,"           '/* 연평균산출세액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax)  hga070t_calc_tax,"                   '/* 산출세액 */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub1 retire_sub,"  '======>/* 퇴직소득세액공제 */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax deci_tax,"  '======>/* 결정세액 */

'명예퇴직 중간정산자 2005-06-30
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.honor_amt) honor_amt,"            '/* 퇴직급여액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub2) + FLOOR(hga070t.special_sub2) h_income_sub,"  '/* 퇴직소득공제 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std2)  h_hga070t_tax_std,"                     '/* 퇴직소득과표 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std2)  h_hga070t_avr_tax_std,"             '/* 연평균과세표준 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax2)  h_hga070t_avr_calc_tax,"           '/* 연평균산출세액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax2)  h_hga070t_calc_tax,"                   '/* 산출세액 */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub2 h_retire_sub,"  '======>/* 퇴직소득세액공제 */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax2 h_deci_tax,"  '======>/* 결정세액 */

'합계 
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.retire_amt) + FLOOR(hga070t.honor_amt)  tot_income,"            '/* 퇴직급여액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub) + FLOOR(hga070t.special_sub) +FLOOR(hga070t.income_sub2) + FLOOR(hga070t.special_sub2) t_income_sub,"  '/* 퇴직소득공제 */
                    
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std) + FLOOR(hga070t.tax_std2)  t_hga070t_tax_std,"                     '/* 퇴직소득과표 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std) + FLOOR(hga070t.avr_tax_std2)  t_hga070t_avr_tax_std,"             '/* 연평균과세표준 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax) + FLOOR(hga070t.avr_calc_tax2)  t_hga070t_avr_calc_tax,"           '/* 연평균산출세액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax) + FLOOR(hga070t.calc_tax2)  t_hga070t_calc_tax,"                   '/* 산출세액 */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub1 + hga070t.tax_sub2 t_retire_sub,"  '======>/* 퇴직소득세액공제 */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax + hga070t.deci_income_tax2 t_deci_tax,"  '======>/* 결정세액 */
'------------------------
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2)  hga070t_deci_income_tax,"       '/* 소득세 */
                    lgStrSQL = lgStrSQL & " 0 deci_farm_tax,"                                    '/* 농특세 */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.deci_res_tax)  hga070t_deci_res_tax,"             '/* 주민세 */
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.deci_income_tax) + FLOOR(hga070t.deci_income_tax2)+ FLOOR(hga070t.deci_res_tax))  deci_sum,"  '/* 계 */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_income_tax), 0)  hfa050t_old_income_tax,"  '/* 소득세 */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_farm_tax), 0)  hfa050t_old_farm_tax,"      '/* 농특세 */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_res_tax), 0)  hfa050t_old_res_tax,"        '/* 주민세 */         
                    lgStrSQL = lgStrSQL & " ISNULL((FLOOR(hfa050t.old_income_tax) + FLOOR(hfa050t.old_farm_tax) + FLOOR(hfa050t.old_res_tax)), 0)  old_sum," '/* 계 */

					lgStrSQL = lgStrSQL & " SPACE(17)  empty_2,"    '/* 공란 */                    
                    lgStrSQL = lgStrSQL & " haa010t.emp_no haa010t_emp_no"
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hga070t,"
                    lgStrSQL = lgStrSQL & " hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hga070t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_yy *= hfa050t.year_yy"
                    lgStrSQL = lgStrSQL & " AND hga070t.emp_no *= hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
'                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hga070t.honor_retire_flag=" & FilterVar("Y", "''", "S")     
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_dt BETWEEN" & pComp & " " &  pCode
'response.Write "lgStrSQL3:" & lgStrSQL                    
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
