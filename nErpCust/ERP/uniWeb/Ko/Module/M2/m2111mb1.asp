<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MB1
'*  4. Program Name         : 구매요청등록 
'*  5. Program Desc         : 구매요청등록 
'*  6. Component List       : PM2G119.cMLookupPurReqS / PM2G111.cMMaintPurReqS / PB3S106.cBLkUpItemByPlt
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : MINHJ
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%	

call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")

	Dim lgOpModeCRUD
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
   
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "changeItemPlant"
             Call ChangeItemPlant()
    End Select
    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iM21119																	'☆ : 입력/수정용 ComProxy Dll 사용 변수															'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strDt
	Dim strDefFrDate
	Dim strDefToDate
	
	Dim I1_m_pur_req 
    Dim E1_b_pur_org 
    Dim E2_b_minor 
    Dim E3_b_plant 
    Dim E4_b_item 
    Dim E5_m_pur_req 
    Dim E6_b_unit_of_measure 
    Dim E7_b_acct_dept 
    Dim E8_b_storage_location 
    Dim E9_b_minor 
    Dim E10_b_pur_grp 
    Dim E11_b_biz_partner 
            
	Const M092_E3_plant_cd = 0
    Const M092_E3_plant_nm = 1

    Const M092_E4_item_cd = 0
    Const M092_E4_item_nm = 1
    Const M092_E4_spec = 2

    Const M092_E5_pr_no = 0
    Const M092_E5_pr_sts = 1
    Const M092_E5_req_qty = 2
    Const M092_E5_req_unit = 3
    Const M092_E5_req_cfm_qty = 4
    Const M092_E5_req_dt = 5
    Const M092_E5_req_prsn = 6
    Const M092_E5_dlvy_dt = 7
    Const M092_E5_sl_cd = 8
    Const M092_E5_pur_plan_dt = 9
    Const M092_E5_mrp_ord_no = 10
    Const M092_E5_ord_qty = 11
    Const M092_E5_rcpt_qty = 12
    Const M092_E5_procure_type = 13
    Const M092_E5_pr_type = 14
    Const M092_E5_mrp_run_no = 15
    Const M092_E5_pur_org = 16
    Const M092_E5_iv_qty = 17
    Const M092_E5_req_dept = 18
    Const M092_E5_sppl_cd = 19
    Const M092_E5_pur_grp = 20
    Const M092_E5_tracking_no = 21
    Const M092_E5_base_req_qty = 22
    Const M092_E5_base_req_unit = 23
    Const M092_E5_so_no = 24
    Const M092_E5_so_seq_no = 25
    Const M092_E5_ext1_cd = 26
    Const M092_E5_ext1_qty = 27
    Const M092_E5_ext1_amt = 28
    Const M092_E5_ext1_rt = 29
    Const M092_E5_ext2_cd = 30
    Const M092_E5_ext2_qty = 31
    Const M092_E5_ext2_amt = 32
    Const M092_E5_ext2_rt = 33
    Const M092_E5_ext3_cd = 34
    Const M092_E5_ext3_qty = 35
    Const M092_E5_ext3_amt = 36
    Const M092_E5_ext3_rt = 37

    On Error Resume Next      
                                                           '☜: Protect system from crashing
    Err.Clear                                              '☜: Clear Error status

	I1_m_pur_req = UCase(Trim(Request("txtReqNo")))
									  
    Set iM21119 = Server.CreateObject("PM2G119.cMLookupPurReqS") 
   
    If CheckSYSTEMError(Err,True) = true Then 		
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

    Call iM21119.M_LOOKUP_PUR_REQ_SVR(gStrGlobalCollection, I1_m_pur_req, E1_b_pur_org, E2_b_minor, _
			E3_b_plant, E4_b_item, E5_m_pur_req, E6_b_unit_of_measure, E7_b_acct_dept, E8_b_storage_location, _
			E9_b_minor, E10_b_pur_grp, E11_b_biz_partner)

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21119 = Nothing
		'-- issue for 8999 by Byun Jee Hyun 2004-11-30
		Response.Write "<Script Language=vbscript>" & vbCr	
		Response.Write "	Call Parent.SetDefaultVal	" & vbCr
		Response.Write "</Script>" & vbCr										'☜: ComProxy Unload
		'-- end of issue for 8999
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

	'-----------------------
	'Result data display area
	'----------------------- 
	strDefFrDate = "1900-01-01"
	strDefToDate = "2999-12-31"
	
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "strDefFrDate = """ & UNIDateClientFormat(strDefFrDate) & """" & vbCr
    Response.Write "strDefToDate = """ & UNIDateClientFormat(strDefToDate) & """" & vbCr
	Response.Write "Parent.frm1.txtReqNo2.value     = """ & ConvSPChars(E5_m_pur_req(M092_E5_pr_no))      & """" & vbCr
	Response.Write "Parent.frm1.txtPlantCd.value    = """ & ConvSPChars(E3_b_plant(M092_E3_plant_cd))   & """" & vbCr
	Response.Write "Parent.frm1.txtPlantNm.value    = """ & ConvSPChars(E3_b_plant(M092_E3_plant_nm))      & """" & vbCr
	Response.Write "Parent.frm1.txtItemCd.value     = """ & ConvSPChars(E4_b_item(M092_E4_item_cd))   & """" & vbCr
	Response.Write "Parent.frm1.txtitemNm.value     = """ & ConvSPChars(E4_b_item(M092_E4_item_nm))      & """" & vbCr
	Response.Write "Parent.frm1.txtSpec.value       = """ & ConvSPChars(E4_b_item(M092_E4_Spec))      & """" & vbCr
	Response.Write "Parent.frm1.txtDlvyDt.text      = """ & UNIDateClientFormat(E5_m_pur_req(M092_E5_dlvy_dt))   & """" & vbCr
	Response.Write "Parent.frm1.txtReqDt.text       = """ & UNIDateClientFormat(E5_m_pur_req(M092_E5_req_dt))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqQty.text      = """ & UNINumClientFormat(E5_m_pur_req(M092_E5_req_qty),ggQty.DecPoint,0)  & """" & vbCr
	Response.Write "Parent.frm1.txtReqUnitCd.value  = """ & ConvSPChars(E5_m_pur_req(M092_E5_req_unit))  & """" & vbCr
	Response.Write "Parent.frm1.txtDeptCd.value     = """ & ConvSPChars(E5_m_pur_req(M092_E5_req_dept))  & """" & vbCr
	Response.Write "Parent.frm1.txtDeptNm.value     = """ & ConvSPChars(E7_b_acct_dept)  & """" & vbCr
	Response.Write "Parent.frm1.txtEmpCd.value      = """ & ConvSPChars(E5_m_pur_req(M092_E5_req_prsn))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageCd.value  = """ & ConvSPChars(E5_m_pur_req(M092_E5_sl_cd))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageNm.value  = """ & ConvSPChars(E8_b_storage_location)  & """" & vbCr
	Response.Write "Parent.frm1.txtTrackingNo.value = """ & ConvSPChars(E5_m_pur_req(M092_E5_tracking_no))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqStateCd.value = """ & ConvSPChars(E5_m_pur_req(M092_E5_pr_sts))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqStateNm.value = """ & ConvSPChars(E9_b_minor)  & """" & vbCr
	Response.Write "Parent.frm1.txtPoQty.text       = """ & UNINumClientFormat(E5_m_pur_req(M092_E5_ord_qty),ggQty.DecPoint,0)  & """" & vbCr
	Response.Write "Parent.frm1.txtGmQty.text       = """ & UNINumClientFormat(E5_m_pur_req(M092_E5_rcpt_qty),ggQty.DecPoint,0)  & """" & vbCr
	Response.Write "Parent.frm1.hdnProcurType.value = """ & ConvSPChars(E5_m_pur_req(M092_E5_procure_type))  & """" & vbCr
	Response.Write "Parent.frm1.hdnMrpNo.value      = """ & ConvSPChars(E5_m_pur_req(M092_E5_mrp_ord_no))  & """" & vbCr
	Response.Write "Parent.frm1.txtOrgCd.value      = """ & ConvSPChars(E5_m_pur_req(M092_E5_pur_org))  & """" & vbCr
	Response.Write "Parent.frm1.txtOrgNm.value      = """ & ConvSPChars(E1_b_pur_org)  & """" & vbCr
	Response.Write "Parent.frm1.txtReqStateCd.value = """ & ConvSPChars(E5_m_pur_req(M092_E5_pr_sts))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqStateNm.value = """ & ConvSPChars(E9_b_minor)  & """" & vbCr
	Response.Write "Parent.frm1.txtReqTypeCd.value  = """ & ConvSPChars(E5_m_pur_req(M092_E5_pr_type))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqTypeNm.value  = """ & ConvSPChars(E2_b_minor)  & """" & vbCr
	
	Response.Write "If """ & Trim(E5_m_pur_req(M092_E5_tracking_no)) & """ = ""*"" Then " & vbCr
	Response.Write "	Parent.frm1.hdnTrackingflg.value= ""N""  " & vbCr
	Response.Write "Else  " & vbCr
	Response.Write "	Parent.frm1.hdnTrackingflg.value= ""Y""  " & vbCr
	Response.Write "End If " & vbCr                          
	
	Response.Write "Parent.DbQueryOk "           & vbCr
    
    Response.Write "If """ & Trim(E5_m_pur_req(M092_E5_pr_sts)) & """ <> ""RQ""  Or """ & Trim(E5_m_pur_req(M092_E5_sppl_cd)) & """ <> """" Then " & vbCr
	Response.Write "	Call Parent.ChangeTag(True,True) "           & vbCr
	Response.Write "	Call Parent.SetFocusToDocument(""M"")  " & vbCr
	Response.Write "	parent.frm1.vspdData2.focus " & vbCr
    Response.Write "Else " & vbCr
	Response.Write "	Call Parent.ChangeTag(False,True) "           & vbCr
    Response.Write "	Call Parent.changeTagTracking() "           & vbCr
	Response.Write "	Call Parent.SetFocusToDocument(""M"")  " & vbCr
	Response.Write "	parent.frm1.txtReqDt.focus " & vbCr
    Response.Write "End If "                                      & vbCr
    Response.Write "</Script>" & vbCr
		 
    Set iM21119 = Nothing															'☜: Unload Comproxy
	Exit sub
End Sub	

'============================================================================================================
' Name : SubBizSave
' Desc : 
'============================================================================================================
Sub SubBizSave()
	Dim iCommandSent
	Dim iM21111
	Dim lgIntFlgMode
	
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

	If Len(Trim(Request("txtDlvyDt"))) Then
		If UNIConvDate(Request("txtDlvyDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation,	"", "",	I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtDlvyDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
    End	If
		
    If Len(Trim(Request("txtReqDt"))) Then
		If UNIConvDate(Request("txtReqDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation,	"", "",	I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtReqDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
    End	If
    
    If Request("txtFlgMode") = "" Then									
		Call DisplayMsgBox("700112", vbInformation,	"", "",	I_MKSCRIPT)
		Exit Sub 
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))							
	
	Set	iM21111 = CreateObject("PM2G111.cMMaintPurReqS")

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21111 = Nothing
		Exit Sub												
	
	End if	
	
    If lgIntFlgMode = OPMD_CMODE Then
    	iCommandSent = "CREATE"
		if Trim(Request("txtReqNo2")) <> "" then
    		I1_m_pur_req(M069_I1_pr_no) = UCase(Trim(Request("txtReqNo2")))
    	End if    	
    ElseIf lgIntFlgMode	= OPMD_UMODE Then
    	I1_m_pur_req(M069_I1_pr_no) 	= UCase(Trim(Request("txtReqNo2")))
		iCommandSent = "UPDATE"
	End	If
    
    I3_b_plant	= UCase(Trim(Request("txtPlantCd")))
	I2_b_item	= UCase(Trim(Request("txtItemCd")))
	I1_m_pur_req(M069_I1_dlvy_dt) 	= UNIConvDate(Request("txtDlvyDt"))
	I1_m_pur_req(M069_I1_req_dt)	= UNIConvDate(Request("txtReqDt"))
	I1_m_pur_req(M069_I1_req_qty) 	= UNIConvNum(Request("txtReqQty"),0)
	I1_m_pur_req(M069_I1_req_unit)	= UCase(Trim(Request("txtReqUnitCd")))
	I1_m_pur_req(M069_I1_req_dept)	= UCase(Trim(Request("txtDeptCd")))
	I1_m_pur_req(M069_I1_req_prsn) 	= Trim(Request("txtEmpCd"))
	I1_m_pur_req(M069_I1_sl_cd) 	= UCase(Trim(Request("txtStorageCd")))
	If Trim(Request("txtTrackingNo")) <> "" Then
		I1_m_pur_req(M069_I1_tracking_no)	= UCase(Trim(Request("txtTrackingNo")))
	Else
		I1_m_pur_req(M069_I1_tracking_no)	= "*"
	End If
	I1_m_pur_req(M069_I1_procure_type)	= UCase(Trim(Request("hdnProcurType")))
	I1_m_pur_req(M069_I1_pr_sts)		= UCase(Trim(Request("txtReqStateCd")))
	I1_m_pur_req(M069_I1_mrp_ord_no)	= UCase(Trim(Request("hdnMrpNo")))
	I1_m_pur_req(M069_I1_pur_org)		= UCase(Trim(Request("txtOrgCd")))
    I1_m_pur_req(M069_I1_pr_type)		= "E"
    
    export_Req_No =  iM21111.M_MAINT_PUR_REQ_SVR(gStrGlobalCollection, iCommandSent, I1_m_pur_req, CStr(I2_b_item), CStr(I3_b_plant))
	
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iM21111 = Nothing									
		Exit Sub												
	End if
	'-----------------------
	'Result	data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write "Parent.frm1.txtReqNo.value    = """ & ConvSPChars(export_Req_No)   & """" & vbCr
	Response.Write "End If " & vbCr
	Response.Write " Parent.DbSaveOk() "           & vbCr
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
    
    iStrPlantCd = UCase(Trim(Request("txtPlantCd")))
    iStrItemCd = UCase(Trim(Request("txtItemCd")))
    
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
	Response.Write "Parent.frm1.txtPlantCd.value	 = """ & ConvSPChars(E8_b_plant(P003_E8_plant_cd))  & """" & vbCr
	Response.Write "Parent.frm1.txtPlantNm.value     = """ & ConvSPChars(E8_b_plant(P003_E8_plant_nm))  & """" & vbCr
	Response.Write "Parent.frm1.txtItemCd.value      = """ & ConvSPChars(E7_b_item(P003_E7_item_cd))  & """" & vbCr
	Response.Write "Parent.frm1.txtItemNm.value      = """ & ConvSPChars(E7_b_item(P003_E7_item_nm))  & """" & vbCr
	Response.Write "Parent.frm1.txtReqUnitCd.value   = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur))  & """" & vbCr
	Response.Write "Parent.frm1.txtOrgCd.value       = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_pur_org))  & """" & vbCr
	Response.Write "Parent.frm1.txtOrgNm.value       = """ & ConvSPChars(E1_b_pur_org(P003_E1_pur_org_nm))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageCd.value   = """ & ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_cd))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageNm.value   = """ & ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_nm))  & """" & vbCr
	Response.Write "Parent.frm1.hdnTrackingflg.value = """ & ConvSPChars(E6_b_item_by_plant(P003_E6_tracking_flg))  & """" & vbCr
	Response.Write "	Call Parent.changeTagTracking() "           & vbCr
	Response.Write "</Script>" & vbCr
End Sub	


%>
