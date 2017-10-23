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
    Call HideStatusWnd                                                              '��: Hide Processing message
    BiDx = 1

    lgErrorStatus     = "NO"													    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '��: "M"(Spread #1) "S"(Spread #2)
    
    strMode      = Request("txtMode")                                               '��: Read Operation Mode (CRUD)
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
                Set CTFnm = Fso.CreateTextFile (FPnm,True)                              'text�� ������ File�� ����            
                
                CTFnm.Write lgstrData                                                   'Text ����κ�                       
                DFnm = Fso.GetFileName(FPnm)            
                CTFnm.close    
                Set CTFnm = nothing
            Else
                Call DisplayMsgBox("700100", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
        strFilePath = strFilePath  & "files/u2000/"  '2002.02.01 /files ���� ���� u2000�� ����:���߿� ������ ����Ǹ� �����ؾ� ��.
	    strFilePath = strFilePath & Request("txtFileName")

End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(CQuery)        
Dim strWhere    
Dim pComp
 
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            pComp = "="
            strWhere = FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & CQuery
            
            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '��: Make sql statements       
            Call SubBizQueryMulti()    
        Case "B"
            pComp = ""
            strWhere = FilterVar(lgKeyStream(5),"NULL", "S") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery
            strWhere = strWhere & " GROUP BY hfa100t.year_area_cd,"                                                       '�ٲ� Query
            strWhere = strWhere & " hfa100t.tax_biz_cd,"
            strWhere = strWhere & " hfa100t.own_rgst_no,"
            strWhere = strWhere & " hfa100t.year_area_nm,"
            strWhere = strWhere & " hfa100t.repre_nm,"
            strWhere = strWhere & " hfa100t.co_own_rgst_no"

            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '��: Make sql statements                   
        Case "C"        
            pComp = ""
            strWhere = replace(FilterVar(lgKeyStream(5),"NULL", "S"),gComDateType,"") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND haa010t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & CQuery

            Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '��: Make sql statements                           
    End Select       
End Sub	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()    

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call ASubBizQueryMulti()        
End Sub    
'============================================================================================================
' Name : ASubBizQueryMulti()
' Desc : Query ASheet Data from Db
'============================================================================================================
Sub ASubBizQueryMulti()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
        
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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status    
    
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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
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
            lgstrData = lgstrData & SetFixSrting(ClgObjRs("haa010t_nat_cd"),"","",2,"")   '2002 ���������ڵ� 
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
            
'������ �߰������� 2005-06-30

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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))                
                Case "A"
                    lgStrSQL = " SELECT " & FilterVar("A", "''", "S") & "  record_type,"	'/* ���ڵ屸�� : A�� ���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"	'/* �ڷᱸ�� : 22���� ���� */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(2), "''", "S") & " dcl_date,"	'/* ���⿬���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & " p_type,"		'/* ������(�븮��)���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & "  mag_no,"	'/* �����븮�ΰ�����ȣ */  
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  "		'/* 2004 hometax id */                                       
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  "						'/* 2004 �������α׷��ڵ� ��Ÿ */                      
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_rgst_no,"			'/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " year_area_nm  biz_area_nm,"			'/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM,  "		'����� �μ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_NAME,  "			'����ڸ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_TEL,  "			'����� ��ȭ��ȣ 2004  

                    lgStrSQL = lgStrSQL & " " & FilterVar("101", "''", "S") & " kr_code,"                                 '/* ����ѱ��ڵ� : 101�� ���� */
'                    lgStrSQL = lgStrSQL & " " & FilterVar(lgKeyStream(1), "''", "S") & " term_code,"                                 '/* ������Ⱓ�ڵ� -> �Էº��� */
                    lgStrSQL = lgStrSQL & " '1'  term_code,"
                    
                    lgStrSQL = lgStrSQL & " SPACE(371) empty"                               '/* ���� */ 
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd" & pComp & " " &  pCode
'Response.Write "lgStrSQL2:" & lgStrSQL  

                Case "B" 
                    lgStrSQL = " SELECT hfa100t.year_area_cd  singo_org_cd,"
                    lgStrSQL = lgStrSQL & " " & FilterVar("B", "''", "S") & "  record_type,"                                   '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"                                    '/* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"               '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"               '/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  com_rgst_no,"                '/* �ֹ�(����)��Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"                     '/* ��ǥ��(����) */
                    lgStrSQL = lgStrSQL & " COUNT(hga070t.emp_no) com_no,"                      '/* ��(��)����Ǽ�(C���ڵ��) */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0", "''", "S") & "  old_com_no,"                                    '/* ��(��)����Ǽ�(D���ڵ��) -> 0���� ������ ���� */ 
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.tot_prov_amt))  tot_prov_amt,"    '/* �ҵ�ݾ� �Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2))  deci_income_tax,"  '/* �ҵ漼���������Ѱ� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000000", "''", "S") & " tot_tax,"                               '/* ���μ����������Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_res_tax)) deci_res_tax,"         '/* �ֹμ����������Ѱ� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000000", "''", "S") & " deci_farm_tax,"                         '/* ��Ư�����������Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2) + FLOOR(hga070t.deci_res_tax)) deci_sum,"  '/* ���������Ѱ� */
                    lgStrSQL = lgStrSQL & " SPACE(1) d_code,"                                   '/* �ڷ�����ڵ� */
                    lgStrSQL = lgStrSQL & " SPACE(352) empty"                                   '/* ���� */
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hga070t,"
                    lgStrSQL = lgStrSQL & " hfa100t"                                              '�ٲ� Query
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
                    lgStrSQL = lgStrSQL & " " & FilterVar("C", "''", "S") & "  record_type,"                                           '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"                                            '/* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"                       '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00", "''", "S") & " old_com_no,"                                           '/* ��(��)�ٹ�ó�� -> ������ '00'���� ���� */
                    lgStrSQL = lgStrSQL & " CASE WHEN (hdf020t.res_flag IS NULL OR hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " ) THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("2", "''", "S") & " END  hdf020t_res_flag,"  '/* �����ڱ����ڵ� */
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd = " & FilterVar("KR", "''", "S") & ") THEN " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE haa010t.nat_cd END  haa010t_nat_cd,"  '2002 ���������ڵ� 
                    
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.entr_dt < convert(datetime," & FilterVar(lgKeyStream(8) & "-01-01", "''", "S")&") THEN " & FilterVar(lgKeyStream(8) & "-01-01" ,"NULL", "S")
                    lgStrSQL = lgStrSQL & " ELSE haa010t.entr_dt END haa010t_entr_dt,"  
                    lgStrSQL = lgStrSQL & " hga070t.RETIRE_DT  hga070t_retire_dt,"  

                    lgStrSQL = lgStrSQL & " haa010t.name  haa010t_name,"                                '/* ���� */  
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.nat_cd = " & FilterVar("KR", "''", "S") & " THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("9", "''", "S") & " END for_type,"                                     '/* ���ܱ��α����ڵ� */
                    lgStrSQL = lgStrSQL & " haa010t.zip_cd zip,"                                        '/* �����ȣ : 2000�� �������� ���� */
                    lgStrSQL = lgStrSQL & " haa010t.addr address,"                                      '/* �ּ� : 2000�� �������� ���� */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"                                    '/* �ֹ�(����)��Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.retire_amt) + FLOOR(hga070t.etc_amt))  retire_amt,"  '/* �����޿� */  
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.honor_amt)  hga070t_honor_amt,"                '/* ���������� �Ǵ� �߰������޿� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.corp_insur)  hga070t_corp_insur,"              '/* ��ü��������� */ 
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.retire_amt) + FLOOR(hga070t.etc_amt)) + FLOOR(hga070t.honor_amt)+FLOOR(hga070t.corp_insur)   hga070t_tot_prov_amt,"          '/* �� */
                    lgStrSQL = lgStrSQL & " hga070t.entr_dt entr_dt,"                                    '/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL & " hga070t.retire_dt retire_dt,"                                '/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetServeMonths(hga070t.entr_dt,hga070t.retire_dt) hga070t_tot_duty_mm," '/* ��(��)�ٹ����ټӿ���  -2002.03.25*/
                    
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & " old_entr_dt,"                                     '======>/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & " old_retire_dt,"                                   '======>/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000", "''", "S") & " old_duty,"                                            '======>/* ��(��)�ٹ��� �ټӿ��� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000", "''", "S") & " d_duty,"                                              '======>/* �ߺ����� */
                    
                    lgStrSQL = lgStrSQL & " ceiling(dbo.ufn_GetServeMonths(hga070t.entr_dt,hga070t.retire_dt) / 12.0) hga070t_duty_cnt,"  '/* �ټӿ��� -2002.03.25 */
                    

'������ �߰������� 2005-06-30
                    lgStrSQL = lgStrSQL & " haa010t.entr_dt h_entr_dt,"                                    '/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL & " haa010t.retire_dt h_retire_dt,"                                '/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL & " hga070t.tot_duty_mm2," '/* ��(��)�ٹ����ټӿ���  -2002.03.25*/

                    
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " h_old_entr_dt,"                                     '======>/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " h_old_retire_dt,"                                   '======>/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " h_old_duty,"                                            '======>/* ��(��)�ٹ��� �ټӿ��� */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " h_d_duty,"                                              '======>/* �ߺ����� */
                    
                    lgStrSQL = lgStrSQL & " ceiling(dbo.ufn_GetServeMonths(haa010t.entr_dt,haa010t.retire_dt) / 12.0) haa010t_duty_cnt,"  '/* �ټӿ��� -2002.03.25 */
'--------------                    
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.retire_amt) retire_tot_prov_amt,"            '/* �����޿��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub) + FLOOR(hga070t.special_sub) income_sub,"  '/* �����ҵ���� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std)  hga070t_tax_std,"                     '/* �����ҵ��ǥ */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std)  hga070t_avr_tax_std,"             '/* ����հ���ǥ�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax)  hga070t_avr_calc_tax,"           '/* ����ջ��⼼�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax)  hga070t_calc_tax,"                   '/* ���⼼�� */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub1 retire_sub,"  '======>/* �����ҵ漼�װ��� */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax deci_tax,"  '======>/* �������� */

'������ �߰������� 2005-06-30
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.honor_amt) honor_amt,"            '/* �����޿��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub2) + FLOOR(hga070t.special_sub2) h_income_sub,"  '/* �����ҵ���� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std2)  h_hga070t_tax_std,"                     '/* �����ҵ��ǥ */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std2)  h_hga070t_avr_tax_std,"             '/* ����հ���ǥ�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax2)  h_hga070t_avr_calc_tax,"           '/* ����ջ��⼼�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax2)  h_hga070t_calc_tax,"                   '/* ���⼼�� */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub2 h_retire_sub,"  '======>/* �����ҵ漼�װ��� */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax2 h_deci_tax,"  '======>/* �������� */

'�հ� 
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.retire_amt) + FLOOR(hga070t.honor_amt)  tot_income,"            '/* �����޿��� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.income_sub) + FLOOR(hga070t.special_sub) +FLOOR(hga070t.income_sub2) + FLOOR(hga070t.special_sub2) t_income_sub,"  '/* �����ҵ���� */
                    
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.tax_std) + FLOOR(hga070t.tax_std2)  t_hga070t_tax_std,"                     '/* �����ҵ��ǥ */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_tax_std) + FLOOR(hga070t.avr_tax_std2)  t_hga070t_avr_tax_std,"             '/* ����հ���ǥ�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.avr_calc_tax) + FLOOR(hga070t.avr_calc_tax2)  t_hga070t_avr_calc_tax,"           '/* ����ջ��⼼�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.calc_tax) + FLOOR(hga070t.calc_tax2)  t_hga070t_calc_tax,"                   '/* ���⼼�� */
                    
                    lgStrSQL = lgStrSQL & " hga070t.tax_sub1 + hga070t.tax_sub2 t_retire_sub,"  '======>/* �����ҵ漼�װ��� */
                    lgStrSQL = lgStrSQL & " hga070t.deci_income_tax + hga070t.deci_income_tax2 t_deci_tax,"  '======>/* �������� */
'------------------------
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.deci_income_tax)+FLOOR(hga070t.deci_income_tax2)  hga070t_deci_income_tax,"       '/* �ҵ漼 */
                    lgStrSQL = lgStrSQL & " 0 deci_farm_tax,"                                    '/* ��Ư�� */
                    lgStrSQL = lgStrSQL & " FLOOR(hga070t.deci_res_tax)  hga070t_deci_res_tax,"             '/* �ֹμ� */
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.deci_income_tax) + FLOOR(hga070t.deci_income_tax2)+ FLOOR(hga070t.deci_res_tax))  deci_sum,"  '/* �� */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_income_tax), 0)  hfa050t_old_income_tax,"  '/* �ҵ漼 */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_farm_tax), 0)  hfa050t_old_farm_tax,"      '/* ��Ư�� */
                    lgStrSQL = lgStrSQL & " ISNULL(FLOOR(hfa050t.old_res_tax), 0)  hfa050t_old_res_tax,"        '/* �ֹμ� */         
                    lgStrSQL = lgStrSQL & " ISNULL((FLOOR(hfa050t.old_income_tax) + FLOOR(hfa050t.old_farm_tax) + FLOOR(hfa050t.old_res_tax)), 0)  old_sum," '/* �� */

					lgStrSQL = lgStrSQL & " SPACE(17)  empty_2,"    '/* ���� */                    
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

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
' Name : SetFixSrting(�Է°�,�񱳹���,��ü����,��������,�������Ĺ���)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '�Է°��� ��������������� �Է°��� ���̸� 0���� �Ѵ�.
        Cnt = 0     
    Else
    
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2                                                  '�ѱۺκи� ���̸� ���� 2���Ѵ�.
            Else
                MCnt = MCnt + 1
            End If
        Next
        Cnt = MCnt
                 
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '�񱳹��ڰ� ������� 
        Else                                                                         '�񱳹��ڰ� �����Ұ�� �񱳹��ڸ� �� �������� �Է°������Ѵ�.
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
    
    If InPos = "" Then                                                              '�������̰� �������� ���� ��� �Է¹��� ���̰� �������̰� �ȴ�.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '���������ϰ��(defalut) �������� ���� ���� ������ ���ڰ� �ԷµǸ� ������ ����(default)�κ��� ��ü���ڷ� ü���.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '���������� 
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
