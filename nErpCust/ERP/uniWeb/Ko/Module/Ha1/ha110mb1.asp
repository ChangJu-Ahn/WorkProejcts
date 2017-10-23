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
	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey,lgStrPrevKey1,lgStrPrevKey2    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")
	
    Dim strEmpno
    Dim strNo
    
    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '��: "M"(Spread #1) "S"(Spread #2)
    If lgCurrentSpd = "A" Then    
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)    
	elseif lgCurrentSpd = "B" Then    
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)
	elseif lgCurrentSpd = "C" Then   
		lgStrPrevKey2 = UNICInt(Trim(Request("lgStrPrevKey2")),0)	
	end if	 	
        
    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()        
Dim strWhere    
Dim pComp
Dim str
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case UCase(Trim(lgCurrentSpd))
        Case "A"
            pComp = "="
            strWhere = FilterVar(lgKeyStream(7), "''", "S")
         
        Case "B"
            pComp = ""
            strWhere = FilterVar(lgKeyStream(5),"NULL", "S") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & " GROUP BY hfa100t.year_area_cd,"                                                       '�ٲ� Query
            strWhere = strWhere & " hfa100t.tax_biz_cd,"
            strWhere = strWhere & " hfa100t.own_rgst_no,"
            strWhere = strWhere & " hfa100t.year_area_nm,"
            strWhere = strWhere & " hfa100t.repre_nm,"
            strWhere = strWhere & " hfa100t.co_own_rgst_no"
          
        Case "C"
            pComp = ""
            strWhere = FilterVar(lgKeyStream(5),"NULL", "S") & " AND " & FilterVar(lgKeyStream(6),"NULL", "S") 
            strWhere = strWhere & " AND haa010t.year_area_cd = " & FilterVar(lgKeyStream(7), "''", "S")
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
                     
    End Select

    Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '��: Make sql statements    
    Call SubBizQueryMulti()    
End Sub	
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("kr_code"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("term_code"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("empty"))
                
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
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("tot_prov_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_income_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("tot_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_res_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_farm_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_sum"), 0,0)
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("empty"))
                Case "C"
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & iDx
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_com_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hdf020t_res_flag"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_nat_cd"))  '2002 ���������ڵ� 
                     
                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("haa010t_entr_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("hga070t_retire_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_name")  )   
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("for_type"))
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("retire_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_honor_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_corp_insur"), 0,0)       
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_tot_prov_amt"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("entr_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("retire_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hga070t_tot_duty_mm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_entr_dt") )    
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_retire_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_duty"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("d_duty"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hga070t_duty_cnt"))

'������ �߰������� 2005-06-30

                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("h_entr_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & UNIConvDateToYYYYMMDD(lgObjRs("h_retire_dt"),gServerDateFormat,"")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tot_duty_mm2"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("h_old_entr_dt") )    
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("h_old_retire_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("h_old_duty"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("h_d_duty"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_duty_cnt"))
                                          
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("retire_tot_prov_amt"), 0,0)        
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("income_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_avr_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_avr_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("retire_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_tax"), 0,0)                                     

                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("honor_amt"), 0,0)        
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_income_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_hga070t_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_hga070t_avr_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_hga070t_avr_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_hga070t_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_retire_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("h_deci_tax"), 0,0) 
                       
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("tot_income"), 0,0)                        
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_income_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_hga070t_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_hga070t_avr_tax_std"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_hga070t_avr_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_hga070t_calc_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_retire_sub"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("t_deci_tax"), 0,0)                                           
'-------------------------------------------------------------------------------------------
          
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_deci_income_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hga070t_deci_res_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_farm_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("deci_sum"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hfa050t_old_income_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hfa050t_old_res_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("hfa050t_old_farm_tax"), 0,0)
                     lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("old_sum"), 0,0)
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
            DFnm = "C:\ea" & li_biz_own_rgst_no       
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
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    
'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)    
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
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
					
                    lgStrSQL = " SELECT  top " & iSelCount & "" & FilterVar("A", "''", "S") & "  record_type,"	'/* ���ڵ屸�� : A�� ���� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"						'/* �ڷᱸ�� : 22���� ���� */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(2), "''", "S") & " dcl_date,"		'/* ���⿬���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & " p_type,"			'/* ������(�븮��)���� -> �Էº��� */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & "  mag_no,"				'/* �����븮�ΰ�����ȣ */
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  "		'/* 2004 hometax id */  
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  "						'/* 2004 �������α׷��ڵ� ��Ÿ */                                                             
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_rgst_no,"   '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm) biz_area_nm,"  '/* ���θ�(��ȣ) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM,  "		'����� �μ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_NAME,  "			'����ڸ� 2004   
                    lgStrSQL = lgStrSQL & " WORKER_TEL,  "			'����� ��ȭ��ȣ 2004  
                    lgStrSQL = lgStrSQL & " " & FilterVar("101", "''", "S") & " kr_code,"                                 '/* ����ѱ��ڵ� : 101�� ���� */
'                    lgStrSQL = lgStrSQL & " " & FilterVar(lgKeyStream(1), "''", "S") & " term_code,"                                 '/* ������Ⱓ�ڵ� -> �Էº��� */
                    lgStrSQL = lgStrSQL & " '1'  term_code,"
                    
                    lgStrSQL = lgStrSQL & " SPACE(191) empty"                               '/* ���� */ 
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd" & pComp & " " & pCode

                Case "B" 
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
                    
                    lgStrSQL = " SELECT  top " & iSelCount & " hfa100t.year_area_cd  singo_org_cd,"
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
                    lgStrSQL = lgStrSQL & " 0 tot_tax,"                               '/* ���μ����������Ѱ� */
                    lgStrSQL = lgStrSQL & " 0 deci_farm_tax,"                         '/* ��Ư�����������Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_res_tax)) deci_res_tax,"         '/* �ֹμ����������Ѱ� */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hga070t.deci_income_tax) +FLOOR(hga070t.deci_income_tax2)+ FLOOR(hga070t.deci_res_tax)) deci_sum,"  '/* ���������Ѱ� */
                    lgStrSQL = lgStrSQL & " SPACE(131) empty"                                   '/* ���� */
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hga070t,"
                    lgStrSQL = lgStrSQL & " hfa100t"                                              '�ٲ� Query
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hga070t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_yy *= hfa050t.year_yy"
                    lgStrSQL = lgStrSQL & " AND hga070t.emp_no *= hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S")  
'                   lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S")
                    lgStrSQL = lgStrSQL & " AND hdf020t.retire_give = " & FilterVar("Y", "''", "S")  
                    lgStrSQL = lgStrSQL & " AND hga070t.honor_retire_flag=" & FilterVar("Y", "''", "S")
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_dt BETWEEN" & pComp & " " &  pCode
'Response.Write lgStrSQL
'Response.End
                Case "C" 
                    call  CommonQueryRs(" RE_TAX_SUB1 , RE_SUB_LIMIT "," hda000t "," 1=1" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
					re_tax_sub1 = Replace(Trim(lgF0), Chr(11), "")
					re_tax_sub1 = CInt(re_tax_sub1) / 100
					re_sub_limit = Replace(Trim(lgF1), Chr(11), "")

                    If re_tax_sub1 = "" OR re_tax_sub1 = "X" Then
						re_tax_sub1 = "0.25"
						re_sub_limit = "120000"
					end if 
					                
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey2 + 1
                    
                    lgStrSQL = " SELECT  top " & iSelCount & " hfa100t.year_area_cd  singo_area_cd," 
                    lgStrSQL = lgStrSQL & " " & FilterVar("C", "''", "S") & "  record_type,"                                           '/* ���ڵ屸�� */
                    lgStrSQL = lgStrSQL & " " & FilterVar("25", "''", "S") & " data_type,"                                            '/* �ڷᱸ�� */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* ������ */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_rgst_no,"                       '/* ����ڵ�Ϲ�ȣ */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00", "''", "S") & " old_com_no,"                                           '/* ��(��)�ٹ�ó�� -> ������ '00'���� ���� */
                    lgStrSQL = lgStrSQL & " CASE WHEN (hdf020t.res_flag IS NULL OR hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " ) THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("2", "''", "S") & " END hdf020t_res_flag,"                             '/* �����ڱ����ڵ� */                    
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd = " & FilterVar("KR", "''", "S") & ") THEN " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE haa010t.nat_cd END  haa010t_nat_cd,"  '2002 ���������ڵ� 
                    
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.entr_dt < convert(datetime," & FilterVar(lgKeyStream(8) & "-01-01", "''", "S")& ") THEN " & FilterVar(lgKeyStream(8) & "-01-01" ,"NULL", "S")
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
                    lgStrSQL = lgStrSQL & " (FLOOR(hga070t.retire_amt) + FLOOR(hga070t.etc_amt)) + FLOOR(hga070t.honor_amt)+FLOOR(hga070t.corp_insur)  hga070t_tot_prov_amt,"          '/* �� */
                    lgStrSQL = lgStrSQL & " hga070t.entr_dt entr_dt,"                                    '/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL & " hga070t.retire_dt retire_dt,"                                '/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetServeMonths(hga070t.entr_dt,hga070t.retire_dt) hga070t_tot_duty_mm," '/* ��(��)�ٹ����ټӿ���  -2002.03.25*/
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " old_entr_dt,"                                     '======>/* ��(��)�ٹ����Ի翬���� */
                    lgStrSQL = lgStrSQL &   FilterVar("00000000", "''", "S") & " old_retire_dt,"                                   '======>/* ��(��)�ٹ�����翬���� */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " old_duty,"                                            '======>/* ��(��)�ٹ��� �ټӿ��� */
                    lgStrSQL = lgStrSQL &   FilterVar("0000", "''", "S") & " d_duty,"                                              '======>/* �ߺ����� */
                    
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
'                   lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.retire_give = " & FilterVar("Y", "''", "S") & " "                    
                    lgStrSQL = lgStrSQL & " AND hga070t.honor_retire_flag=" & FilterVar("Y", "''", "S")                    
                    lgStrSQL = lgStrSQL & " AND hga070t.retire_dt BETWEEN" & pComp &  " " & pCode
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
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
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
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
