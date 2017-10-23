<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!--

'**********************************************************************************************
'*  1. Module Name	    : Sale,Production
'*  2. Function	Name	    : Sales Order,....
'*  3. Program ID	    : m2111mb1
'*  4. Program Name	    : 구매요청등록멀티 
'*  5. Program Desc	    :
'*  6. Comproxy	List	    : 
'*  7. Modified	date(First) : 1999/09/10
'*  8. Modified	date(Last)  : 1999/09/10
'*  9. Modifier	(First)	    : Mr  Kim
'* 10. Modifier	(Last)	    : MINHJ
'* 11. Comment		    :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*			      this mark(⊙) Means that "may  change"
'*			      this mark(☆) Means that "must change"
'* 13. History		    : 2003/09/24 kimjihyun
'* 14. Business	Logic of m2111ma1(구매요청등록)
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%	

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
    Dim lgPageNo
    Dim lgNextKey
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
  
    lgLngMaxRow	      = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount	      = 100                                 '☜: Fetch count at a time for VspdData
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)	   
    lgNextKey 	     = Trim(Request("lgNextKey"))

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "changeItemPlant"
             Call ChangeItemPlant()
    End Select
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim lgNextKey
	
	'조회조건 
    Dim strPntCd, strItemCd, strDlvyFrDt, strDlvyToDt, strReqFrDt, strReqToDt, strOrgCd, strPrno, strMRP
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    strPntCd	= Trim(Request("txtPlantCd"))		'공장코드 
    strItemCd	= Trim(Request("txtItemCd"))		'품목코드 
    strOrgCd	= Trim(Request("txtORGCd"))			'구매조직코드 
    'strPrno	= Trim(Request("txtReqNo"))			'요청번호 
    strMRP		= Trim(Request("txtMRP"))			'MRP Run 번호 

    If Request("txtDlvyFrDt") = "" Then
    	strDlvyFrDt		= "1900-01-01"
    Else
    	strDlvyFrDt 	= CStr(UNIConvDate(Request("txtDlvyFrDt")))
    End if
    
    If Request("txtDlvyToDt") = "" Then
    	strDlvyToDt 	= "2999-12-31"
    Else
    	strDlvyToDt 	= CStr(UNIConvDate(Request("txtDlvyToDt")))
    End if 
    
    
    If Request("txtReqFrDt") = "" Then
    	strReqFrDt 		= "1900-01-01"
    Else
    	strReqFrDt 		= CStr(UNIConvDate(Request("txtReqFrDt")))
    End if
    
    If Request("txtReqToDt") = "" Then
    	strReqToDt 		= "2999-12-31"
    Else
    	strReqToDt 		= CStr(UNIConvDate(Request("txtReqToDt")))
    End if 
    
    if Len(strPntCd) then
		iKey1 = " and a.plant_cd =  " & FilterVar(strPntCd, "''", "S") & " "
    End if
    
    if Len(strItemCd) then
		iKey1 = iKey1 & " and a.item_cd =  " & FilterVar(strItemCd, "''", "S") & " "
    End if
    
    if Len(strOrgCd) then
		iKey1 = iKey1 & " and a.pur_org =  " & FilterVar(strOrgCd, "''", "S") & " "
    End if
      
    iKey1 = iKey1 & " and a.dlvy_dt >=  " & FilterVar(strDlvyFrDt , "''", "S") & ""
    iKey1 = iKey1 & " and a.dlvy_dt <=  " & FilterVar(strDlvyToDt , "''", "S") & ""
    iKey1 = iKey1 & " and a.req_dt >=  " & FilterVar(strReqFrDt , "''", "S") & ""
    iKey1 = iKey1 & " and a.req_dt <=  " & FilterVar(strReqToDt , "''", "S") & ""
    
	'MRP Run 번호 추가 (2005.12.14)
    If Len(Trim(strMRP)) then
		iKey1 = iKey1 & " AND A.MRP_RUN_NO = " & FilterVar(strMRP, "''", "S") & " "
    End if

    if Len(Trim(Request("lgNextKey"))) then
		iKey1 = iKey1 & " and a.pr_no <=  " & FilterVar(Request("lgNextKey"), "''", "S") & " "
    End if

    iKey1 = iKey1 & " Order by a.pr_no Desc"

    Call SubMakeSQLStatements("Q",iKey1)                                 '☆ : Make sql statements
    Call SubOpenDB(lgObjConn)

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgPageNo= ""
		lgNextKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

'    	Call SubSkipRs(lgObjRs,lgMaxCount * lgPageNo)

        lgstrData = ""
        iDx       = 0
        
        Do While Not lgObjRs.EOF
            iDx =  iDx + 1
                    
      if idx =< lgMaxCount then
        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pr_no"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("plant_cd"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(ConvSPChars(lgObjRs("req_qty")),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("req_unit"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(ConvSPChars(lgObjRs("dlvy_dt")))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(ConvSPChars(lgObjRs("req_dt")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pur_org"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("req_dept"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("req_prsn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sl_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tracking_no"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(ConvSPChars(lgObjRs("pur_plan_dt")))
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pr_sts"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sts_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pr_type"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("type_nm"))
		if ConvSPChars(lgObjRs("tracking_no")) <> "*" Then
			lgstrData = lgstrData & Chr(11) & "Y"
		Else
			lgstrData = lgstrData & Chr(11) & "N"
		End if
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("procure_type"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mrp_ord_no"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sppl_cd"))
           '------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
    end if
    

               If iDx > lgMaxCount Then
               lgPageNo = lgPageNo + 1
		       lgNextKey = ConvSPChars(lgObjRs("pr_no"))
	
				   Exit Do
			   End If   
			    
	    lgObjRs.MoveNext
	    
        Loop 
    End If
    
    If iDx <= lgMaxCount Then
        lgPageNo =  ""
    End If   

    Call SubHandleError("Q",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                           '☜: Release RecordSSet
    '-----------------------
    'Result data display area
    '----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " With Parent" & vbCr
	Response.Write "    if Trim(""" & lgErrorStatus & """)= ""NO"" Then " & vbCr
	Response.Write "	 .ggoSpread.Source  	 = .frm1.vspdData" & vbCr
	Response.Write "         .ggoSpread.SSShowData """ & lgstrData & """" & vbCr
	Response.Write "         .lgPageNo		 = """ & lgPageNo & """" & vbCr
	Response.Write "	 .lgNextKey 	 = """ & lgNextKey & """" & vbCr
	Response.Write "        .DBQueryOk" & vbCr
	Response.Write "    End if" & vbCr
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr


End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode)
    Dim iSelCount
	'call svrmsgbox(pComp , vbinformation, i_mkscript)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
  
    Select Case (pDataType)
        Case "Q"
	   lgStrSQL = "Select Top 101 a.pr_no, a.plant_cd, c.plant_nm, a.item_cd, b.item_nm, b.spec, a.req_qty, a.req_unit, a.dlvy_dt, a.req_dt,a.pur_org, a.req_dept, a.req_prsn, a.sl_cd,a.tracking_no," & _
		      "a.sppl_cd, a.pur_grp,a.pur_plan_dt, a.pr_sts,dbo.ufn_GetCodeName(" & FilterVar("M2101", "''", "S") & ",a.pr_Sts) sts_nm,a.pr_type, dbo.ufn_GetCodeName(" & FilterVar("M2102", "''", "S") & ", a.pr_type) type_nm,a.procure_type, a.mrp_ord_no,a.sppl_cd"
           lgStrSQL = lgStrSQL & " From  m_pur_req a, b_item b, b_plant c  "
           lgStrSQL = lgStrSQL & " Where a.plant_cd = c.plant_cd and a.item_cd = b.item_cd " & pCode
    End Select             

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
        Case "Q"
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
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
	Err.Clear													    '☜: Protect system	from crashing
	
	Dim iM21111
	Dim lgIntFlgMode
	Dim LngMaxRow
	
	Dim iStrSpread
	Dim iErrorPosition
	Dim export_Req_No            
	
	'lgIntFlgMode = CInt(Request("txtFlgMode"))							

	Dim SpdCount
	Dim i
	
	'LngMaxRow = CLng(Request("txtMaxRows"))		

	iStrSpread = Trim(Request("txtSpread"))
	SpdCount = CInt(Request("SpdCount"))

	For i = 1 to SpdCount
		iStrSpread = iStrSpread & Request("txtSpread" & i)
	Next

	Set	iM21111 = CreateObject("PM2G111.cMMaintPurReqMultiS")


    	If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21111 = Nothing
		Exit Sub												
	End if	
 
        Call iM21111.M_MAINT_PUR_REQ_MULTI_SVR(gStrGlobalCollection, iStrSpread, iErrorPosition)
	'call svrmsgbox(iErrorPosition , vbinformation, i_mkscript)
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
		Set iM21111 = Nothing									
		Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
		Exit Sub
	End If
	'-----------------------
	'Result	data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	'Response.Write " If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	'Response.Write "Parent.frm1.txtReqNo.value    = """ & ConvSPChars(export_Req_No)   & """" & vbCr
	'Response.Write "End If " & vbCr
	Response.Write " Parent.Dbsaveok() "           & vbCr
	Response.Write "</Script>" & vbCr

	Set	iM21111	= Nothing

End Sub	
'============================================================================================================
' Name : SubBizDelete
' Desc : 
'============================================================================================================
Sub SubBizDelete()
    
	Dim iM21111
	Dim iCommandSent

	Dim I1_m_pur_req
    Dim I2_b_item
    Dim I3_b_plant
	Dim export_Req_No           
	            
	Const M069_I1_req_qty = 0
    Const M069_I1_req_dt = 1
    Const M069_I1_req_prsn = 2
    Const M069_I1_dlvy_dt = 3
    Const M069_I1_sl_cd = 4
    Const M069_I1_pr_no = 5
    Const M069_I1_pr_type = 6
    Const M069_I1_req_unit = 7
    Const M069_I1_req_dept = 8
    Const M069_I1_tracking_no = 9
    Const M069_I1_procure_type = 10
    Const M069_I1_pr_sts = 11
    Const M069_I1_pur_plan_dt = 12
    Const M069_I1_sppl_cd = 13
    Const M069_I1_pur_org = 14
    Const M069_I1_pur_grp = 15
    Const M069_I1_mrp_ord_no = 16
    Const M069_I1_mrp_run_no = 17
    Const M069_I1_so_no = 18
    Const M069_I1_so_seq_no = 19
    Const M069_I1_ext1_cd = 20
    Const M069_I1_ext1_qty = 21
    Const M069_I1_ext1_amt = 22
    Const M069_I1_ext1_rt = 23
    Const M069_I1_ext2_cd = 24
    Const M069_I1_ext2_qty = 25
    Const M069_I1_ext2_amt = 26
    Const M069_I1_ext2_rt = 27
    Const M069_I1_ext3_cd = 28
    Const M069_I1_ext3_qty = 29
    Const M069_I1_ext3_amt = 30
    Const M069_I1_ext3_rt = 31
    
    ReDim I1_m_pur_req(M069_I1_ext3_rt)
    
	On Error Resume Next
	
	Err.Clear													    '☜: Protect system	from crashing
					
	I1_m_pur_req(M069_I1_pr_no) = Trim(Request("txtReqNo2"))

    Set	iM21111 = CreateObject("PM2G111.cMMaintPurReqS")

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21111 = Nothing
		Exit Sub												
	
	End if
	
	iCommandSent = "Delete"
	
	export_Req_No = iM21111.M_MAINT_PUR_REQ_SVR(gStrGlobalCollection, iCommandSent, I1_m_pur_req, CStr(I2_b_item), CStr(I3_b_plant))

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21111 = Nothing									
		Exit Sub												
	
	End if
	
	'-----------------------
	'Result	data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "	Call Parent.DbDeleteOk() "           & vbCr
	Response.Write "</Script>" & vbCr
	
	Set	iM21111	= Nothing						    '☜: Unload	Comproxy
	
		
End Sub	


'============================================================================================================
' Name : ChangeItemPlant
' Desc : 
'============================================================================================================
Sub ChangeItemPlant()
    
    Dim I1_plant_cd
    Dim I2_b_item
    Dim E1_b_pur_org
    Dim E2_b_item_group
    Dim E3_for_issued_b_storage_location
    Dim E4_for_major_b_storage_location
    Dim E5_i_material_valuation
    Dim E6_b_item_by_plant
    Dim E7_b_item
    Dim E8_b_plant
    
    ' E1_b_pur_org
    Const P003_E1_pur_org = 0
    Const P003_E1_pur_org_nm = 1
    Const P003_E1_valid_fr_dt = 2
    Const P003_E1_valid_to_dt = 3
    Const P003_E1_usage_flg = 4

    ' E2_b_item_group
    Const P003_E2_item_group_cd = 0
    Const P003_E2_item_group_nm = 1
    Const P003_E2_leaf_flg = 2

    ' E3_for_issued b_storage_location
    Const P003_E3_sl_cd = 0
    Const P003_E3_sl_type = 1
    Const P003_E3_sl_nm = 2

    ' E4_for_major b_storage_location
    Const P003_E4_sl_cd = 0
    Const P003_E4_sl_type = 1
    Const P003_E4_sl_nm = 2

    ' E5_i_material_valuation
    Const P003_E5_prc_ctrl_indctr = 0
    Const P003_E5_moving_avg_prc = 1
    Const P003_E5_std_prc = 2
    Const P003_E5_prev_std_prc = 3

    ' E6_b_item_by_plant
    Const P003_E6_procur_type = 0
    Const P003_E6_order_unit_mfg = 1
    Const P003_E6_order_lt_mfg = 2
    Const P003_E6_order_lt_pur = 3
    Const P003_E6_order_type = 4
    Const P003_E6_order_rule = 5
    Const P003_E6_req_round_flg = 6
    Const P003_E6_fixed_mrp_qty = 7
    Const P003_E6_min_mrp_qty = 8
    Const P003_E6_max_mrp_qty = 9
    Const P003_E6_round_qty = 10
    Const P003_E6_round_perd = 11
    Const P003_E6_scrap_rate_mfg = 12
    Const P003_E6_ss_qty = 13
    Const P003_E6_prod_env = 14
    Const P003_E6_mps_flg = 15
    Const P003_E6_issue_mthd = 16
    Const P003_E6_mrp_mgr = 17
    Const P003_E6_inv_check_flg = 18
    Const P003_E6_lot_flg = 19
    Const P003_E6_cycle_cnt_perd = 20
    Const P003_E6_inv_mgr = 21
    Const P003_E6_major_sl_cd = 22
    Const P003_E6_abc_flg = 23
    Const P003_E6_mps_mgr = 24
    Const P003_E6_recv_inspec_flg = 25
    Const P003_E6_inspec_lt_mfg = 26
    Const P003_E6_inspec_mgr = 27
    Const P003_E6_valid_from_dt = 28
    Const P003_E6_valid_to_dt = 29
    Const P003_E6_item_acct = 30
    Const P003_E6_single_rout_flg = 31
    Const P003_E6_prod_mgr = 32
    Const P003_E6_issued_sl_cd = 33
    Const P003_E6_issued_unit = 34
    Const P003_E6_order_unit_pur = 35
    Const P003_E6_var_lt = 36
    Const P003_E6_scrap_rate_pur = 37
    Const P003_E6_pur_org = 38
    Const P003_E6_prod_inspec_flg = 39
    Const P003_E6_final_inspec_flg = 40
    Const P003_E6_ship_inspec_flg = 41
    Const P003_E6_inspec_lt_pur = 42
    Const P003_E6_option_flg = 43
    Const P003_E6_over_rcpt_flg = 44
    Const P003_E6_over_rcpt_rate = 45
    Const P003_E6_damper_flg = 46
    Const P003_E6_damper_max = 47
    Const P003_E6_damper_min = 48
    Const P003_E6_reorder_pnt = 49
    Const P003_E6_std_time = 50
    Const P003_E6_add_sel_rule = 51
    Const P003_E6_add_sel_value = 52
    Const P003_E6_add_seq_rule = 53
    Const P003_E6_add_seq_atrid = 54
    Const P003_E6_rem_sel_rule = 55
    Const P003_E6_rem_sel_value = 56
    Const P003_E6_rem_seq_rule = 57
    Const P003_E6_rem_seq_atrid = 58
    Const P003_E6_llc = 59
    Const P003_E6_tracking_flg = 60
    Const P003_E6_valid_flg = 61
    Const P003_E6_work_center = 62
    Const P003_E6_order_from = 63
    Const P003_E6_cal_type = 64
    Const P003_E6_line_no = 65
    Const P003_E6_atp_lt = 66
    Const P003_E6_etc_flg1 = 67
    Const P003_E6_etc_flg2 = 68

    ' E7_b_item
    Const P003_E7_item_cd = 0
    Const P003_E7_item_nm = 1
    Const P003_E7_formal_nm = 2
    Const P003_E7_spec = 3
    Const P003_E7_item_acct = 4
    Const P003_E7_item_class = 5
    Const P003_E7_hs_cd = 6
    Const P003_E7_hs_unit = 7
    Const P003_E7_unit_weight = 8
    Const P003_E7_unit_of_weight = 9
    Const P003_E7_basic_unit = 10
    Const P003_E7_draw_no = 11
    Const P003_E7_item_image_flg = 12
    Const P003_E7_phantom_flg = 13
    Const P003_E7_blanket_pur_flg = 14
    Const P003_E7_base_item_cd = 15
    Const P003_E7_proportion_rate = 16
    Const P003_E7_valid_flg = 17
    Const P003_E7_valid_from_dt = 18
    Const P003_E7_valid_to_dt = 19

    ' E8_b_plant
    Const P003_E8_plant_cd = 0
    Const P003_E8_plant_nm = 1
    
    Dim iStrPlantCd
    Dim iStrItemCd
    Dim iStrRow
            
    On Error Resume Next                       '☜: Protect system from crashing
    
    Err.Clear								    '☜: Protect system	from crashing
	
	Dim iB1b119
	
    Set	iB1b119 = CreateObject("PB3S106.cBLkUpItemByPlt")

    '-----------------------
    'Com action	result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iB1b119 = Nothing									
		Exit Sub												
	
	End if
    
    iStrPlantCd = FilterVar(UCase(Trim(Request("txtPlantCd"))),"","SNM")
    iStrItemCd = FilterVar(UCase(Trim(Request("txtItemCd"))),"","SNM")
	iStrRow = Request("txtRow")
    Call iB1b119.B_LOOK_UP_ITEM_BY_PLANT(gStrGlobalCollection, iStrPlantCd, iStrItemCd, E1_b_pur_org, _
												E2_b_item_group, E3_for_issued_b_storage_location, _
												E4_for_major_b_storage_location, _
												E5_i_material_valuation, _
												E6_b_item_by_plant, _
												E7_b_item, _
												E8_b_plant)
            
            
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iB1b119 = Nothing									
		Exit Sub												
	
	End if

	'-----------------------
	'Result	data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " With Parent.frm1.vspdData"      		& vbCr
	Response.Write " 	.Row  	=  " & iStrRow    			& vbCr
	Response.Write " 	.Col 	= Parent.C_PlantCd "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E8_b_plant(P003_E8_plant_cd)) & """" 	& vbCr	
	Response.Write " 	.Col 	= Parent.C_ItemCd  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_item_cd)) & """" 	& vbCr	
	Response.Write " 	.Col 	= Parent.C_ItemNm  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E7_b_item(P003_E7_item_nm)) & """" 	& vbCr	
	Response.Write " 	.Col 	= Parent.C_ReqUnit	  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur)) & """" 	& vbCr	
	Response.Write " 	.Col 	= Parent.C_PurOrg	  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_pur_org)) & """" & vbCr	
	Response.Write " 	.Col 	= Parent.C_StorageCd 	  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_cd)) & """" & vbCr	
	Response.Write " 	.Col 	= Parent.C_HdnTrackingflg  	  "       	& vbCr
	Response.Write " 	.text   = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_tracking_flg)) & """" & vbCr	
		
	Response.Write "	Call Parent.changeTagTracking() "           & vbCr
	Response.Write " End With" & vbCr
	Response.Write "</Script>" & vbCr
	 
	
End Sub	


%>
