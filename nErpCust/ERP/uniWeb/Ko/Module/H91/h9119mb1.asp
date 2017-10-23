<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
	Const C_SHEETMAXROWS_D = 100
 	Dim lgStrPrevKey,lgStrPrevKey1,lgStrPrevKey2,lgStrPrevKey3    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                   '☜: Clear Error status
	  
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
	elseif lgCurrentSpd = "D" Then    
		lgStrPrevKey3 = UNICInt(Trim(Request("lgStrPrevKey3")),0)	
	end if	 	    
    
	strEmpno      = Split(Request("C_EMP_NO_2"),gColSep)   
	strNo      = Split(Request("C_NO2"),gColSep)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
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
            pComp = "="
            strWhere = FilterVar(lgKeyStream(6), "''", "S")
        Case "B"
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year," & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '바뀐 Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
            strWhere = strWhere & " GROUP BY hfa100t.year_area_cd,"                                                       '바뀐 Query
            strWhere = strWhere & " hfa100t.tax_biz_cd,"
            strWhere = strWhere & " hfa100t.own_rgst_no,"
            strWhere = strWhere & " hfa100t.year_area_nm,"
            strWhere = strWhere & " hfa100t.repre_nm,"
            strWhere = strWhere & " hfa100t.co_own_rgst_no,"
            strWhere = strWhere & " hfa100t.addr"
        Case "C"   
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year, " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '바뀐 Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
        Case "D"
            pComp = ">"
            strWhere = " CONVERT(VARCHAR(4), DATEPART(year, " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & "))"        
            strWhere = strWhere & " OR (DATEPART(year, haa010t.retire_dt) = CONVERT(NUMERIC(4), " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & ") AND haa010t.retire_resn = " & FilterVar("6", "''", "S") & "))"
            strWhere = strWhere & " AND haa010t.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
            strWhere = strWhere & " AND hfa050t.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")    
            strWhere = strWhere & " AND hfa100t.year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S")                '바뀐 Query
            strWhere = strWhere & " AND haa010t.year_area_cd = hfa100t.year_area_cd"
    End Select
    
    Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements    
    Call SubBizQueryMulti()    
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
    Dim DFnm        
    Dim li_biz_own_rgst_no
    Dim Oldres_no,Cwork_no
    Dim i,strDNO
    Dim c_per_sub, c_spouse_sub, c_fam_sub, c_old_sub, c_paria_sub, c_lady_sub, c_chl_rear_sub ,c_old_sub2
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""    
        lgStrPrevKey1 = ""    
        lgStrPrevKey2 = ""    
        lgStrPrevKey3 = ""    
'       Call SetErrorStatus()
    Else    
        Select Case UCase(Trim(lgCurrentSpd))
             Case "A"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)    
    		 Case "B"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)    
    		 Case "C"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey2)    
    		 Case "D"
    				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey3)    
        End Select    		    		 
    
        lgstrData = ""
        Oldres_no = ""
        Cwork_no = 0
        li_biz_own_rgst_no = Trim(lgKeyStream(4))        
        iDx = 1
        
         Call CommonQueryRs(" old_sub2 ","HFA020T"," 1=1 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        c_old_sub2     = Replace(lgF0, Chr(11), "")
        Call CommonQueryRs("per_sub, spouse_sub, fam_sub, old_sub, paria_sub, lady_sub, chl_rear_sub ","HFA020T"," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        c_per_sub      = Replace(lgF0, Chr(11), "") 
        c_spouse_sub   = Replace(lgF1, Chr(11), "")
        c_fam_sub      = Replace(lgF2, Chr(11), "")
        c_old_sub      = Replace(lgF3, Chr(11), "")
        c_paria_sub    = Replace(lgF4, Chr(11), "")
        c_lady_sub     = Replace(lgF5, Chr(11), "")
        c_chl_rear_sub = Replace(lgF6, Chr(11), "")                                   
        
        Do While Not lgObjRs.EOF
            Select Case UCase(Trim(lgCurrentSpd))
                 Case "A"
                     If Trim(li_biz_own_rgst_no) = "" Or Trim(li_biz_own_rgst_no) <> Trim(replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")) Then 
                         li_biz_own_rgst_no = Trim(replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-",""))
                         li_biz_own_rgst_no = Left(li_biz_own_rgst_no,7) & "." & Right(li_biz_own_rgst_no,3)
                     End If

                     Call CommonQueryRs("count(*) ","HFA100T","year_area_cd = " & FilterVar(lgKeyStream(6), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mag_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hometax_id"))  
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("taxpgm_cd"))                     
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_DEPT_NM"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_NAME"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WORKER_TEL"))
					 lgstrData = lgstrData & Chr(11) & ConvSPChars(Replace(lgF0, Chr(11), ""))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("kr_code"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_gigan"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("empty"))
                Case "B"
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & iDx
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("repre_nm"))
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("co_own_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("com_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_com_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tot"))
                     lgstrData = lgstrData & Chr(11) & lgObjRs("dec_income_tax")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("tot_tax")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("dec_res_tax")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("dec_farm_tax")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("dec_tot")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("empty"))
                Case "C"
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(iDx +lgStrPrevKey2 *100)
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("old_com_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hdf020t_res_flag"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_nat_cd"))  '2002 거주지국코드 
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FOREIGN_SEPARATE_TAX_YN"))  '2004 외국인단일세율 
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_entr_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_retire_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_name"))       
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("for_type"))
                     lgstrData = lgstrData & Chr(11) & left(replace(ConvSPChars(lgObjRs("res_no")),"-",""),13)
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("start_dt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("end_dt"))
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_new_pay_tot")         
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_new_bonus_tot")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa030t_after_bonus_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("new_tot")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_non_tax5")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_non_tax1")        
                     lgstrData = lgstrData & Chr(11) & lgObjRs("non_tax")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("non_tax_sum")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_income_tot_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_income_sub_amt")        
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_income_amt")
      '               lgstrData = lgstrData & Chr(11) & c_per_sub                     
   
'                     IF lgObjRs("hfa050t_spouse_sub_amt") > 0 Then
 '                       lgstrData = lgstrData & Chr(11) & c_spouse_sub                     
  '                   Else
   '                     lgstrData = lgstrData & Chr(11) & 0                    
    '                 End if                     
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_per_sub_amt")   
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_spouse_sub_amt")   
                                         
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_supp_cnt")
                     lgstrData = lgstrData & Chr(11) & CInt(lgObjRs("hfa050t_supp_cnt")) * c_fam_sub
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_old_cnt")
                     lgstrData = lgstrData & Chr(11) & CInt(lgObjRs("old_cnt1")) * c_old_sub + CInt(lgObjRs("old_cnt2")) * c_old_sub2
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_paria_cnt")
                     lgstrData = lgstrData & Chr(11) & CInt(lgObjRs("hfa050t_paria_cnt")) * c_paria_sub
                     IF lgObjRs("hfa050t_lady_sub_amt") > 0 Then
                        lgstrData = lgstrData & Chr(11) & c_lady_sub                     
                     Else
                        lgstrData = lgstrData & Chr(11) & 0                    
                     End if                     
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_chl_rear")
                     lgstrData = lgstrData & Chr(11) & CInt(lgObjRs("hfa050t_chl_rear")) * c_chl_rear_sub

                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_small_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_national_pension_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_insur_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_med_sub_amt")   
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_edu_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_house_fund_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_contr_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_ceremony_amt") '2004
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_std_sub_tot_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_std_sub_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_sub_income_amt")    
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_indiv_anu_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_indiv_anu2_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_invest_sub_sum_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_card_sub_sum_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_our_stock_amt") '2002 우리사주조합출연금 추가 
                     lgstrData = lgstrData & Chr(11) & lgObjRs("special_tax_sum")  '2004
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_tax_std_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_calu_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_income_redu_amt") '2002 세액감면:소득세법 
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_taxes_redu_amt")  
                     lgstrData = lgstrData & Chr(11) & lgObjRs("empty")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_redu_sum_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_income_tax_sub_amt") '세액공제:근로소득 
                     lgstrData = lgstrData & Chr(11) & lgObjRs("empty_4")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_house_repay_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_fore_pay_amt")   
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_poli_tax_sub")  '2004
                     lgstrData = lgstrData & Chr(11) & lgObjRs("empty_3")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_tax_sub_amt")  '2002 세액공제계 
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_dec_income_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_dec_res_tax_amt")   
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_dec_farm_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("dec_tot")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_old_income_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_old_res_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_old_farm_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("old_tot")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_new_income_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_new_res_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("hfa050t_new_farm_tax_amt")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("new_sum_tot")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("empty_2")
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("haa010t_emp_no"))
               Case "D"  
                                
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))                       
                     if UBound(strEmpno) > 0 Then
                        For i = 0 To (Cdbl(UBound(strEmpno)) - 1)
                            If Trim(strEmpno(i)) = Trim(ConvSPChars(lgObjRs("haa010t_emp_no"))) Then                            
                                strDNO = Trim(strNo(i))                            
                            End If
                        Next
	                 Else
	                    strDNO = ""                        
                     End If
                     lgstrData = lgstrData & Chr(11) & Trim(strDNO)
                     lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")
                     lgstrData = lgstrData & Chr(11) & lgObjRs("empty")
                     If Trim(Oldres_no) = "" Then
                        Oldres_no = ConvSPChars(lgObjRs("res_no"))
                        Cwork_no = 1
                     Else
                        If Oldres_no = ConvSPChars(lgObjRs("res_no")) Then                            
                            Cwork_no = Cwork_no + 1
                        Else
                            Oldres_no = ConvSPChars(lgObjRs("res_no"))
                            Cwork_no = 1
                        End If
                     End If
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(left(lgObjRs("res_no"),13))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_nm"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_no"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hfa040t_pay_tot"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bonus_tot"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bonus_amt"))
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_tot"))
                     lgstrData = lgstrData & Chr(11) & Cwork_no
                     lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("empty_2")  )                   

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
				    Case "D"
						lgStrPrevKey3 = lgStrPrevKey3 + 1
				End Select					
               Exit Do
            End If                       
                       
        Loop         
        If Trim(lgCurrentSpd) = "A" then
            DFnm = "C:\c" & li_biz_own_rgst_no       
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
		    Case "D"
		       lgStrPrevKey3 = ""
		End Select		       
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
    Dim Submit_dt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Submit_dt = UniConvDateToYYYYMMDD(lgKeyStream(2),gDateFormat,"")
    
    Select Case Mid(pDataType,2,1)
        Case "R"
            Select Case UCase(Trim(lgCurrentSpd))
                Case "A"
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1                
                    lgStrSQL = " SELECT top " & iSelCount & " " & FilterVar("A", "''", "S") & "   record_type,"              '/* 레코드구분 : A로 고정 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"           '/* 자료구분 : 20으로 고정 */
                    lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(Submit_dt, "''", "S") & "  present_dt,"      '/* 제출연월일 -> 입력변수 */
                    lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & "  present_type,"'/* 제출자(대리인)구분 -> 입력변수 */
                    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3),"''", "S") & "  mag_no," '/* 세무대리인관리번호 */ 
                    lgStrSQL = lgStrSQL & " ISNULL(HOMETAX_ID, ' ') hometax_id,  " '/* 2004 hometax id */ 
                    lgStrSQL = lgStrSQL & " '9000' taxpgm_cd,  " 						 '/* 2004 세무프로그램코드 기타 */                 
                    lgStrSQL = lgStrSQL & " own_rgst_no  biz_own_rgst_no,"   '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm)  biz_area_nm,"  '/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " WORKER_DEPT_NM,  " '담당자 부서 2004   
                    lgStrSQL = lgStrSQL & " WORKER_NAME,  " '담당자명 2004   
                    lgStrSQL = lgStrSQL & " WORKER_TEL,  " '담당자 전화번호 2004                    
'                   lgStrSQL = lgStrSQL & " co_own_rgst_no,"     '/* 주민(법인)등록번호 */
 '                   lgStrSQL = lgStrSQL & " repre_nm,"           '/* 대표자(성명) */
  '                  lgStrSQL = lgStrSQL & " addr  address,"      '/* 주소 -> 2000년 연말정산 폐지 */
   '                 lgStrSQL = lgStrSQL & " tel_no,"             '/* 전화번호 */
                    lgStrSQL = lgStrSQL & " '101'  kr_code,"               '/* 사용한글코드 : 101로 고정 */
                    lgStrSQL = lgStrSQL & " '1'  present_gigan,"           '/* 제출대상기간코드 -> 입력변수 */
  '                  lgStrSQL = lgStrSQL & " SPACE(1)  mod_code,"           '/* 자료수정코드 */
                    lgStrSQL = lgStrSQL & " SPACE(551)  empty"             '/* 공란 */
                    lgStrSQL = lgStrSQL & " FROM hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE year_area_cd" & pComp & pCode
                Case "B"                
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1                
                    lgStrSQL = " SELECT  top " & iSelCount & " hfa100t.year_area_cd  year_area_cd,"           '/* 신고사업장 */========='바뀐 Query
                    lgStrSQL = lgStrSQL & " " & FilterVar("B", "''", "S") & "   record_type,"                          '/* 레코드구분 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"                           '/* 자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.year_area_nm  biz_area_nm,"  '/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " hfa100t.repre_nm  repre_nm,"             '/* 대표자(성명) */
                    lgStrSQL = lgStrSQL & " hfa100t.co_own_rgst_no  co_own_rgst_no,"     '/* 주민(법인)등록번호 */
                    lgStrSQL = lgStrSQL & " hfa100t.addr  address,"                  '/* 주소 -> 2000년 연말정산 폐지 */
                    lgStrSQL = lgStrSQL & " COUNT(hfa050t.emp_no) com_no,"              '/* 주(현)근무처(레코드) 수 */
                    lgStrSQL = lgStrSQL & " SUM(t.emp_no_cnt)  old_com_no,"          '/* 종(전)근무처(레코드) 수*/========='바뀐 Query
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.income_tot_amt))  tot,"   '/* 소득금액총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_income_tax))  dec_income_tax,"  '/* 소득세 결정세액 총계 */
                    lgStrSQL = lgStrSQL & " 0 tot_tax,"                                  '/* 법인세 결정세액 총계 -> 0 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_res_tax))  dec_res_tax,"        '/* 주민세 결정세액 총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_farm_tax))  dec_farm_tax,"      '/* 농특세 결정세액 총계 */
                    lgStrSQL = lgStrSQL & " SUM(FLOOR(hfa050t.dec_income_tax) + FLOOR(hfa050t.dec_farm_tax) + FLOOR(dec_res_tax))  dec_tot,"  '/* 결정세액 총계 */
                    lgStrSQL = lgStrSQL & " SPACE(471) empty"                             '/* 공란 */                                        
                    lgStrSQL = lgStrSQL & " FROM haa010t," 
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hfa100t,"                                              '바뀐 Query
                    lgStrSQL = lgStrSQL & " (SELECT emp_no, COUNT(*) emp_no_cnt"                    '바뀐 Query
                    lgStrSQL = lgStrSQL & " FROM hfa040t"                                           '바뀐 Query
                    lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")                                 '바뀐 Query
                    lgStrSQL = lgStrSQL & " GROUP BY emp_no) AS t"                                  '바뀐 Query
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= t.emp_no"                         '바뀐 Query
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"              '/* 퇴직자 포함 */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode
                Case "C"                
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey2 + 1                
                    lgStrSQL = " SELECT  top " & iSelCount & " " & FilterVar("C", "''", "S") & "   record_type,"                  '/* 레코드구분 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"               ' /* 자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " ISNULL(before.emp_before_count,0)  old_com_no,"  '/* 종(전)근무처(레코드) 수*/
                    lgStrSQL = lgStrSQL & " CASE WHEN (hdf020t.res_flag IS NULL OR hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " ) THEN " & FilterVar("1", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("2", "''", "S") & " END  hdf020t_res_flag,"  '/*  */
                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd = " & FilterVar("KR", "''", "S") & ") THEN " & FilterVar("KR", "''", "S") & ""  'lsn
                    lgStrSQL = lgStrSQL & " ELSE haa010t.nat_cd END  haa010t_nat_cd,"  '2002 거주지국코드 
                    lgStrSQL = lgStrSQL & " CASE WHEN hdf020t.FOREIGN_SEPARATE_TAX_YN ='Y' THEN 1 ELSE 2 END FOREIGN_SEPARATE_TAX_YN,"   '/* 외국인단일세율 */ 2004
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.entr_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") 
                    lgStrSQL = lgStrSQL & " + " & FilterVar("0101", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE CONVERT(VARCHAR(8), haa010t.entr_dt, 112) END  haa010t_entr_dt,"  '/*  */
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.retire_dt IS NULL THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN haa010t.retire_dt > " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
                    lgStrSQL = lgStrSQL & " + " & FilterVar("1231", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " ELSE CONVERT(VARCHAR(8), haa010t.retire_dt, 112) END  haa010t_retire_dt,"  '/*  */ 
                    lgStrSQL = lgStrSQL & " haa010t.name  haa010t_name,"  '/* 이름 */
                    lgStrSQL = lgStrSQL & " CASE WHEN haa010t.nat_cd = " & FilterVar("KR", "''", "S") & " THEN 1 ELSE 9 END  for_type,"  '/*주민번호  */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"  '/* 주민등록번호 */
                    lgStrSQL = lgStrSQL & " haa010t.zip_cd,"          '/* 2000년 연말정산 폐지 */
                    lgStrSQL = lgStrSQL & " haa010t.addr,"            '/* 2000년 연말정산 폐지 */
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.entr_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(2),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.entr_dt >= " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("0101", "''", "S") & " THEN CONVERT(VARCHAR(8), haa010t.entr_dt, 112)"
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("00000000", "''", "S") & " END  start_dt,"  '/* 감면기간시작연월일 */ 
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NULL THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NOT NULL"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt >= " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & " THEN " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(2),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " WHEN ISNULL(haa010t.nat_cd, " & FilterVar("KR", "''", "S") & ") <> " & FilterVar("KR", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND hfa050t.redu_sum <> 0"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt IS NOT NULL"
                    lgStrSQL = lgStrSQL & " AND haa010t.retire_dt < " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("1231", "''", "S") & " THEN CONVERT(VARCHAR(8), haa010t.retire_dt, 112)"
                    lgStrSQL = lgStrSQL & " ELSE " & FilterVar("00000000", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " END  end_dt,"  '/* 감면기간종료연월일 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_pay_tot)  hfa050t_new_pay_tot,"  '/* */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_bonus_tot)  hfa050t_new_bonus_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa030t.after_bonus_amt)  hfa030t_after_bonus_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.new_pay_tot) + FLOOR(hfa050t.new_bonus_tot) + FLOOR(hfa030t.after_bonus_amt)) new_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.non_tax5)  hfa050t_non_tax5,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.non_tax1)  hfa050t_non_tax1,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.non_tax2) + FLOOR(hfa050t.non_tax3) + FLOOR(hfa050t.non_tax4) + FLOOR(hfa050t.non_tax6))  non_tax,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.non_tax1) + FLOOR(hfa050t.non_tax2) + FLOOR(hfa050t.non_tax3) + FLOOR(hfa050t.non_tax4) + FLOOR(hfa050t.non_tax5) + FLOOR(hfa050t.non_tax6))  non_tax_sum,"  '/*  */                     
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_tot_amt)  hfa050t_income_tot_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_sub)  hfa050t_income_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_amt)  hfa050t_income_amt,"  '/*  */
' 2001 연말정산                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.per_sub)  hfa050t_per_sub_amt,"  '/* 본인공제 */
                    'lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(hfa050t.spouse," & FilterVar("N", "''", "S") & " ) = " & FilterVar("Y", "''", "S") & "  THEN 1000000"
'                    lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_spouse_sub_amt,"  '/* 배우자공제 */                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.SPOUSE_SUB)  hfa050t_spouse_sub_amt,"  '/* 배우자공제 */

                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.supp_cnt)  hfa050t_supp_cnt,"  '/* 부양자 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_cnt+ hfa050t.old_cnt2)  hfa050t_old_cnt,hfa050t.old_cnt old_cnt1, hfa050t.old_cnt2 old_cnt2,"  '/* 경로우대 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.paria_cnt)  hfa050t_paria_cnt,"  '/* 장애자 */
                    lgStrSQL = lgStrSQL & " CASE WHEN ISNULL(hfa050t.lady," & FilterVar("N", "''", "S") & " ) = " & FilterVar("Y", "''", "S") & "  THEN 500000"
                    lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_lady_sub_amt,"  '/* 부녀자 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.chl_rear)  hfa050t_chl_rear,"  '/* 자녀양육 */

                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.small_sub)  hfa050t_small_sub_amt,"  '/* 소수추가 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.insur_sub)  hfa050t_insur_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.med_sub)  hfa050t_med_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.edu_sub)  hfa050t_edu_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.house_fund)  hfa050t_house_fund_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.contr_sub)  hfa050t_contr_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.ceremony_amt)  hfa050t_ceremony_amt,"  '2004

                    lgStrSQL = lgStrSQL & " CASE WHEN FLOOR(hfa050t.std_sub) <= 600000 and (hdf020t.FOREIGN_SEPARATE_TAX_YN <> " & FilterVar("Y", "''", "S") & ") THEN 0"
                    lgStrSQL = lgStrSQL & " ELSE FLOOR(hfa050t.std_sub) END  hfa050t_std_sub_tot_amt,"  '/* 계(특별공제) */
                    lgStrSQL = lgStrSQL & " CASE WHEN FLOOR(hfa050t.std_sub) <= 600000 and (hdf020t.FOREIGN_SEPARATE_TAX_YN <> " & FilterVar("Y", "''", "S") & ") THEN 600000"
                    lgStrSQL = lgStrSQL & " ELSE 0 END  hfa050t_std_sub_amt,"     '/* 표준공제 */
                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.national_pension_sub_amt)  hfa050t_national_pension_sub_amt,"  '/* 국민연금 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.sub_income_amt)  hfa050t_sub_income_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu)  hfa050t_indiv_anu_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu2)  hfa050t_indiv_anu2_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.invest_sub_sum)  hfa050t_invest_sub_sum_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.card_sub_sum)  hfa050t_card_sub_sum_amt,"  '/*  */

                    lgStrSQL = lgStrSQL & " CASE WHEN (haa010t.nat_cd <> " & FilterVar("KR", "''", "S") & ") THEN FLOOR(hfa050t.fore_edu_sub_amt) "
                    lgStrSQL = lgStrSQL & " ELSE FLOOR(hfa050t.our_stock_amt)  END  hfa050t_our_stock_amt," '2003 (내국인: 우리사주조합출연금) & (외국인 :외국인교육비임차료)
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.indiv_anu) +FLOOR(hfa050t.indiv_anu2)+FLOOR(hfa050t.invest_sub_sum)+FLOOR(hfa050t.card_sub_sum)+FLOOR(hfa050t.our_stock_amt) special_tax_sum, " '2004
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.tax_std)  hfa050t_tax_std_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.calu_tax)  hfa050t_calu_tax_amt," '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_tax_sub)  hfa050t_income_tax_sub_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " " & FilterVar("00000000", "''", "S") & "  empty_4,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.house_repay)  hfa050t_house_repay_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.stock_save)  hfa050t_stock_save_amt,"  '/* 근로자주식저축 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.fore_pay)  hfa050t_fore_pay_amt,"  '/*  */
 '                   lgStrSQL = lgStrSQL & " FLOOR(hfa050t.long_stock_save_amt)  hfa050t_long_stock_save_amt,"  '/* 증권저축 */
                    lgStrSQL = lgStrSQL & " hfa050t.POLI_TAX_SUB hfa050t_poli_tax_sub ,"	'2004
                    lgStrSQL = lgStrSQL & " SPACE(8) empty_3,"
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.tax_sub_sum)  hfa050t_tax_sub_amt,"  '/* 세액공제계 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.income_redu)  hfa050t_income_redu_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.taxes_redu)  hfa050t_taxes_redu_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " " & FilterVar("0000000000", "''", "S") & " empty,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.redu_sum)  hfa050t_redu_sum_amt,"  '/* 세액감면계 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_income_tax)  hfa050t_dec_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_res_tax)  hfa050t_dec_res_tax_amt,"  '/*  */                    
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.dec_farm_tax)  hfa050t_dec_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.dec_income_tax) + FLOOR(hfa050t.dec_farm_tax) + FLOOR(hfa050t.dec_res_tax))  dec_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_income_tax)  hfa050t_old_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_res_tax)  hfa050t_old_res_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.old_farm_tax)  hfa050t_old_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.old_income_tax) + FLOOR(hfa050t.old_farm_tax) + FLOOR(hfa050t.old_res_tax))  old_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_income_tax)  hfa050t_new_income_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_res_tax)  hfa050t_new_res_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa050t.new_farm_tax)  hfa050t_new_farm_tax_amt,"  '/*  */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa050t.new_income_tax) + FLOOR(hfa050t.new_farm_tax) + FLOOR(hfa050t.new_res_tax))  new_sum_tot,"  '/*  */
                    lgStrSQL = lgStrSQL & " SPACE(4)  empty_2,"    '/* 공란 */
                    lgStrSQL = lgStrSQL & " haa010t.emp_no  haa010t_emp_no" '/* 사번 */                '바뀐 Query
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa030t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hfa100t,"
                    lgStrSQL = lgStrSQL & " (SELECT emp_no, COUNT(*) emp_before_count"
                    lgStrSQL = lgStrSQL & " FROM hfa040t"                    
                    lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
                    lgStrSQL = lgStrSQL & " GROUP BY emp_no) AS before,"
                    lgStrSQL = lgStrSQL & " (SELECT a.emp_no, a.pay_yymm, a.pay_tot_amt"
                    lgStrSQL = lgStrSQL & " FROM hdf070t a"
                    lgStrSQL = lgStrSQL & " WHERE a.pay_yymm = (SELECT MAX(b.pay_yymm)"
                    lgStrSQL = lgStrSQL & " FROM hdf070t b"
                    lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no"  
                    lgStrSQL = lgStrSQL & " AND b.pay_yymm LIKE " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S") & " + " & FilterVar("%", "''", "S") & ""
                    lgStrSQL = lgStrSQL & " AND b.prov_type = " & FilterVar("1", "''", "S") & " "  '/* 월정액급여 */
                    lgStrSQL = lgStrSQL & " AND b.pay_tot_amt <> 0)"
                    lgStrSQL = lgStrSQL & " AND a.prov_type = " & FilterVar("1", "''", "S") & " ) AS t"                    
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hdf020t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= t.emp_no"                                 '바뀐 Query
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= before.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.year_yy *= hfa030t.yy"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no *= hfa030t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"  '/* 퇴직자 포함 */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode                    
                Case "D"
					iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey3 + 1                
                    lgStrSQL = " SELECT  top " & iSelCount & " " & FilterVar("D", "''", "S") & "   record_type,"                                  '/* 레코드구분 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("20", "''", "S") & "  data_type,"                               '/* 자료구분 */
                    lgStrSQL = lgStrSQL & " hfa100t.tax_biz_cd  tax,"           '/* 세무서 */
                    lgStrSQL = lgStrSQL & " hfa100t.own_rgst_no  biz_own_rgst_no,"   '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " SPACE(50)  empty,"                              '/* 공란 */
                    lgStrSQL = lgStrSQL & " haa010t.res_no  res_no,"                        '/* 소득자주민등록번호 */
                    lgStrSQL = lgStrSQL & " hfa040t.a_comp_nm  comp_nm,"                    '/* 법인명(상호) */
                    lgStrSQL = lgStrSQL & " hfa040t.a_comp_no  comp_no,"                    '/* 사업자등록번호 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_pay_tot_amt)  hfa040t_pay_tot,"         '/* 급여총액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_bonus_tot_amt)  bonus_tot,"     '/* 상여총액 */
                    lgStrSQL = lgStrSQL & " FLOOR(hfa040t.a_after_bonus_amt) bonus_amt,"    '/* 인정상여 */
                    lgStrSQL = lgStrSQL & " (FLOOR(hfa040t.a_pay_tot_amt) + FLOOR(hfa040t.a_bonus_tot_amt) + FLOOR(hfa040t.a_after_bonus_amt)) pay_tot," '/* 계 */
                    lgStrSQL = lgStrSQL & " " & FilterVar("", "''", "S") & "  work_no,"                                  '/* work_no : 종전근무처일련번호 */
                    lgStrSQL = lgStrSQL & " SPACE(472)  empty_2,"                            '/* 공란 */
                    lgStrSQL = lgStrSQL & " haa010t.emp_no  haa010t_emp_no" '/* 사번 */                '바뀐 Query                    
                    lgStrSQL = lgStrSQL & " FROM haa010t,"
                    lgStrSQL = lgStrSQL & " hfa040t,"
                    lgStrSQL = lgStrSQL & " hfa050t,"
                    lgStrSQL = lgStrSQL & " hdf020t,"
                    lgStrSQL = lgStrSQL & " hfa100t"
                    lgStrSQL = lgStrSQL & " WHERE haa010t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hfa050t.year_yy = hfa040t.year_yy"
                    lgStrSQL = lgStrSQL & " AND hfa050t.emp_no = hfa040t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.emp_no = hfa050t.emp_no"
                    lgStrSQL = lgStrSQL & " AND hdf020t.res_flag = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND hdf020t.year_mon_give = " & FilterVar("Y", "''", "S") & " "
                    lgStrSQL = lgStrSQL & " AND (haa010t.retire_dt IS NULL"                 '/* 퇴직자 포함 */
                    lgStrSQL = lgStrSQL & " OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt))" & pComp & pCode
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
						if .topleftOK then
							.DBQueryOk
						else
							.lgCurrentSpd = "D"						
							.DBQuery
						end if
                        
                    Case "D"
                        .ggoSpread.Source     = .frm1.vspdData3
                        .lgStrPrevKey3    = "<%=lgStrPrevKey3%>"
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
