<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : B1263MB1
'*  4. Program Name         : 사업자이력등록   
'*  5. Program Desc         : 사업자이력등록   
'*  6. Comproxy List        : PB5CS41.dll, PB5CS44.dll, PB5CS45.dll
'*  7. Modified date(First) : 2000/04/28
'*  8. Modified date(Last)  : 2002/05
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : SeoJinKyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 20021223 - 컬럼추가 사내외구분 으로인해 배열순서 고침  강준구 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
	Err.Clear
	On Error Resume Next                                                             '☜: Protect system from crashing 

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call HideStatusWnd                                                               '☜: Hide Processing message
	lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)


	Select Case lgOpModeCRUD
		Case CStr(UID_M0001)                                                         '☜: Query   
			Call SubBizQuery()
		Case CStr(UID_M0002)                                                         '☜: Save,Update, 
			Call SubBizSave()
		Case CStr(UID_M0003)                '☜:  Delete
			Call SubBizDelete()
		Case "BizPartLookUp"
			Call SubBizQueryBizPartner()
	End Select
    
'============================================================================================================
' Name : SubBizQuery
' Desc :  
'============================================================================================================
Sub SubBizQuery() 

	Dim pB1h031 

	Dim imp_b_biz_partner

	Dim imp_b_biz_partner_history_valid_from_dt

	Dim exp_b_biz_partner
	Const C_exp_b_biz_partner_bp_nm = 0 

	Dim exp_b_biz_partner_history
	Const C_exp_b_biz_partner_history_bp_rgst_no = 0
	Const C_exp_b_biz_partner_history_bp_full_nm = 1
	Const C_exp_b_biz_partner_history_bp_nm = 2
	Const C_exp_b_biz_partner_history_bp_eng_nm = 3
	Const C_exp_b_biz_partner_history_repre_nm = 4
	Const C_exp_b_biz_partner_history_repre_rgst_no = 5
	Const C_exp_b_biz_partner_history_ind_type = 6
	Const C_exp_b_biz_partner_history_ind_class = 7
	Const C_exp_b_biz_partner_history_valid_from_dt = 8
	Const C_exp_b_biz_partner_history_chg_reason = 9
	Const C_exp_b_biz_partner_history_insrt_user_id = 10
	Const C_exp_b_biz_partner_history_insrt_dt = 11
	Const C_exp_b_biz_partner_history_updt_user_id = 12
	Const C_exp_b_biz_partner_history_updt_dt = 13
	Const C_exp_b_biz_partner_history_ext1_qty = 14
	Const C_exp_b_biz_partner_history_ext2_qty = 15
	Const C_exp_b_biz_partner_history_ext1_amt = 16
	Const C_exp_b_biz_partner_history_ext2_amt = 17
	Const C_exp_b_biz_partner_history_ext1_cd = 18
	Const C_exp_b_biz_partner_history_ext2_cd = 19

	Const C_exp_b_biz_partner_history_zip_cd = 20
	Const C_exp_b_biz_partner_history_addr1 = 21
	Const C_exp_b_biz_partner_history_addr2 = 22
	Const C_exp_b_biz_partner_history_addr1_eng = 23
	Const C_exp_b_biz_partner_history_addr2_eng = 24
	Const C_exp_b_biz_partner_history_addr3_eng = 25

    Const C_exp_b_biz_partner_history_email = 26
    Const C_exp_b_biz_partner_history_sub_biz_area = 27
    Const C_exp_b_biz_partner_history_sub_biz_desc = 28

	Dim exp_ind_type_nm
	Const C_exp_ind_class_nm_b_minor_minor_nm = 0

	Dim exp_ind_class_nm
	Const C_exp_ind_type_nm_b_minor_minor_nm = 0


	On Error Resume Next
	Err.Clear                


	imp_b_biz_partner = Trim(Request("txtConBp_cd")) 
	imp_b_biz_partner_history_valid_from_dt= UNIConvDate(Request("txtConValidFromDt"))


	Set pB1h031 = Server.CreateObject("PB5CS45.CbLkBizPartnerHist")
 
	if CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	end If 

	call pB1h031.B_LOOKUP_BIZ_PARTNER_HISTORY(gStrGlobalCollection, imp_b_biz_partner,imp_b_biz_partner_history_valid_from_dt , _
	exp_b_biz_partner_history, exp_b_biz_partner,  exp_ind_type_nm ,exp_ind_class_nm)        


	if CheckSYSTEMError(Err,True) = True Then   
		Response.Write "<Script language=vbs> " & vbCr       
		Response.Write " Parent.frm1.txtConBp_nm.value       = """ & ConvSPChars(exp_b_biz_partner(C_exp_b_biz_partner_bp_nm)) & """" & vbCr 
		Response.Write "</Script> "    & vbCr      
		Set pB1h031 = Nothing  
		Exit Sub
	end If

	Response.Write "<Script language=vbs> " & vbCr       
	Response.Write " Parent.frm1.txtConBp_nm.value      = """ & ConvSPChars(exp_b_biz_partner(C_exp_b_biz_partner_bp_nm)) & """" & vbCr 
	Response.Write " Parent.frm1.txtBp_cd.value         = """ & ConvSPChars(Trim(Request("txtConBp_cd"))) & """" & vbCr       
	Response.Write " Parent.frm1.txtBp_nm.value         = """ & ConvSPChars(exp_b_biz_partner(C_exp_b_biz_partner_bp_nm)) & """" & vbCr        
	Response.Write " Parent.frm1.txtValidFromDt.text    = """ & UNIDateClientFormat(ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_valid_from_dt)))  & """" & vbCr 

	Response.Write " Parent.frm1.txtBp_Rgst_No.value	= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_bp_rgst_no)) & """" & vbCr       
	Response.Write " Parent.frm1.txtCust_full_nm.value	= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_bp_full_nm)) & """" & vbCr        
	Response.Write " Parent.frm1.txtCust_nm.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_bp_nm))      & """" & vbCr 
	Response.Write " Parent.frm1.txtCust_eng_nm.value	= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_bp_eng_nm))  & """" & vbCr       
	Response.Write " Parent.frm1.txtRepre_nm.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_repre_nm))   & """" & vbCr    

	Response.Write " Parent.frm1.txtRepre_Rgst.value	= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_repre_rgst_no)) & """" & vbCr 

	Response.Write " Parent.frm1.txtInd_Class.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_ind_class))      & """" & vbCr       
	Response.Write " Parent.frm1.txtInd_ClassNm.value	= """ & ConvSPChars(exp_ind_class_nm(C_exp_ind_class_nm_b_minor_minor_nm))                 & """" & vbCr            

	Response.Write " Parent.frm1.txtInd_Type.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_ind_type))     & """" & vbCr     
	Response.Write " Parent.frm1.txtInd_TypeNm.value	= """ & ConvSPChars(exp_ind_type_nm(C_exp_ind_class_nm_b_minor_minor_nm))                & """" & vbCr 

	Response.Write " Parent.frm1.txtZIP_cd.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_zip_cd))     & """" & vbCr     
	Response.Write " Parent.frm1.txtADDR1.value			= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_addr1))     & """" & vbCr     
	Response.Write " Parent.frm1.txtADDR2.value			= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_addr2))     & """" & vbCr     
	Response.Write " Parent.frm1.txtADDR1_Eng.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_addr1_eng))     & """" & vbCr
	Response.Write " Parent.frm1.txtADDR2_Eng.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_addr2_eng))     & """" & vbCr     
	Response.Write " Parent.frm1.txtADDR3_Eng.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_addr3_eng))     & """" & vbCr                  

	Response.Write " Parent.frm1.txtChgCause.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_chg_reason))                & """" & vbCr            

	Response.Write " Parent.frm1.txtEMail.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_email))                & """" & vbCr            
    Response.Write " Parent.frm1.txtSubBizArea.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_sub_biz_area))                & """" & vbCr            
    Response.Write " Parent.frm1.txtSubBizDesc.value		= """ & ConvSPChars(exp_b_biz_partner_history(C_exp_b_biz_partner_history_sub_biz_desc))                & """" & vbCr            

	Response.Write " Parent.DbQueryOk "                                                & vbCr   
	Response.Write "</Script> "                       & vbCr      
    
	Set pB1h031 = Nothing     
End Sub    

	
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear 

	Dim exp_b_biz_partner
	Dim exp_b_biz_partner_history
	Dim exp_ind_type_nm
	Dim exp_ind_class_nm

	Dim imp_b_biz_partner
	Dim imp_b_biz_partner_history

	Const C_imp_b_biz_partner_history_valid_from_dt = 0 
	Const C_Imp_contry_cd = 1

	Dim pB1h031
	Dim iCommandSent

	'☜: Clear Error status
	ReDim imp_b_biz_partner_history(1)

	iCommandSent = "DELETE" 

	imp_b_biz_partner = Trim(Request("txtBp_cd"))
	imp_b_biz_partner_history(C_imp_b_biz_partner_history_valid_from_dt)= UNIConvDate(Request("txtValidFromDt"))       
	imp_b_biz_partner_history(C_imp_Contry_cd) = Trim(Request("txtContry_cd"))

	Set pB1h031 = Server.CreateObject("PB5CS44.CbHBizPartnerHisSvr") 

	If CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	End If

	Call pB1h031.B_MAINT_BIZ_PARTNER_HIS_SVR(gStrGlobalCollection, iCommandSent ,imp_b_biz_partner ,imp_b_biz_partner_history)

	If CheckSYSTEMError(Err,True) = True Then 
		Set pB1h031 = Nothing
		Exit Sub
	End If 

	Set pB1h031 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr       
	Response.Write " Parent.DbDeleteOk "    & vbCr   
	Response.Write "</Script> "             & vbCr      
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()     

	Dim imp_b_biz_partner

	Dim imp_b_biz_partner_history(26)

	Dim imp_b_biz_partner_history1

	Const C_bp_rgst_no = 0
	Const C_bp_full_nm = 1
	Const C_bp_nm = 2
	Const C_bp_eng_nm = 3
	Const C_repre_nm = 4
	Const C_repre_rgst_no = 5
	Const C_ind_type = 6
	Const C_ind_class = 7
	Const C_valid_from_dt = 8
	Const C_chg_reason = 9
	Const C_ext1_qty = 10
	Const C_ext2_qty = 11
	Const C_ext1_amt = 12
	Const C_ext2_amt = 13
	Const C_ext1_cd = 14
	Const C_ext2_cd = 15

	Const C_zip_cd = 16
	Const C_addr1 = 17
	Const C_addr2 = 18
	Const C_addr1_eng = 19
	Const C_addr2_eng = 20
	Const C_addr3_eng = 21
	Const C_txtContry_cd = 22    

    Const C_txtEMail = 23    
    Const C_txtSub_biz_area = 24    
    Const C_txtSub_biz_desc = 25    

	Dim pB1h031
	Dim CommandSent
	'global ?
	Dim lgIntFlgMode  


	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status


	lgIntFlgMode = CInt(Request("txtFlgMode"))         '☜: 저장시 Create/Update 판별 

	If lgIntFlgMode = OPMD_CMODE Then
		CommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		CommandSent = "UPDATE"  
	End If

	imp_b_biz_partner= Trim(Request("txtBp_cd"))
	imp_b_biz_partner_history(C_valid_from_dt) = UNIConvDate(Request("txtValidFromDt")) 

	imp_b_biz_partner_history(C_bp_rgst_no) = Trim(Request("txtBp_Rgst_No"))
	imp_b_biz_partner_history(C_bp_full_nm) = Trim(Request("txtCust_full_nm"))
	imp_b_biz_partner_history(C_bp_nm) = Trim(Request("txtCust_nm"))
	imp_b_biz_partner_history(C_bp_eng_nm) = Trim(Request("txtCust_eng_nm"))
	imp_b_biz_partner_history(C_repre_nm) = Trim(Request("txtRepre_nm"))                   
	imp_b_biz_partner_history(C_repre_rgst_no) = Trim(Request("txtRepre_Rgst"))

	imp_b_biz_partner_history(C_ind_type) = Trim(Request("txtInd_Type"))    
	imp_b_biz_partner_history(C_ind_class) = Trim(Request("txtInd_Class"))               
	imp_b_biz_partner_history(C_chg_reason) = Trim(Request("txtChgCause"))

	imp_b_biz_partner_history(C_zip_cd) = Trim(Request("txtZIP_cd"))
	imp_b_biz_partner_history(C_addr1) = Trim(Request("txtADDR1"))
	imp_b_biz_partner_history(C_addr2) = Trim(Request("txtADDR2"))
	imp_b_biz_partner_history(C_addr1_eng) = Trim(Request("txtADDR1_Eng"))
	imp_b_biz_partner_history(C_addr2_eng) = Trim(Request("txtADDR2_Eng"))
	imp_b_biz_partner_history(C_addr3_eng) = Trim(Request("txtADDR3_Eng"))
	imp_b_biz_partner_history(C_txtContry_cd) = Trim(Request("txtContry_cd"))

    imp_b_biz_partner_history(C_txtEMail) = Trim(Request("txtEMail"))
    imp_b_biz_partner_history(C_txtSub_biz_area) = Trim(Request("txtSubBizArea"))
    imp_b_biz_partner_history(C_txtSub_biz_desc) = Trim(Request("txtSubBizDesc"))

	imp_b_biz_partner_history1 = imp_b_biz_partner_history

	Set pB1h031 = Server.CreateObject("PB5CS44.CbHBizPartnerHisSvr") 

	If CheckSYSTEMError(Err,True) = True Then 
	Exit Sub
	End If

	Call pB1h031.B_MAINT_BIZ_PARTNER_HIS_SVR(gStrGlobalCollection, CommandSent ,imp_b_biz_partner ,imp_b_biz_partner_history1)

	If CheckSYSTEMError(Err,True) = True Then 
		Set pB1h031 = Nothing 
		Exit Sub
	End If 	
	Set pB1h031 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr       
	Response.Write " Parent.DbSaveOk "      & vbCr   
	Response.Write "</Script> "    & vbCr      
End Sub

'============================================================================================================
' Name : SubBizQueryBizPartner
' Desc : Save Data 
'============================================================================================================
Sub SubBizQueryBizPartner()

	Dim I1_b_biz_partner
	Dim E1_b_biz_partner

	Const S074_E1_bp_cd = 0
	Const S074_E1_bp_type = 1
	Const S074_E1_bp_rgst_no = 2
	Const S074_E1_bp_full_nm = 3
	Const S074_E1_bp_nm = 4
	Const S074_E1_bp_eng_nm = 5
	Const S074_E1_repre_nm = 6
	Const S074_E1_repre_rgst_no = 7
	Const S074_E1_fnd_dt = 8
	Const S074_E1_zip_cd = 9
	Const S074_E1_addr1 = 10
	Const S074_E1_addr1_eng = 11
	Const S074_E1_ind_type = 12
	Const S074_E1_ind_class = 13
	Const S074_E1_trade_rgst_no = 14
	Const S074_E1_contry_cd = 15
	Const S074_E1_province_cd = 16
	Const S074_E1_currency = 17
	Const S074_E1_tel_no1 = 18
	Const S074_E1_tel_no2 = 19
	Const S074_E1_fax_no = 20
	Const S074_E1_home_url = 21
	Const S074_E1_usage_flag = 22
	Const S074_E1_bp_prsn_nm = 23
	Const S074_E1_bp_contact_pt = 24
	Const S074_E1_biz_prsn = 25
	Const S074_E1_biz_grp = 26
	Const S074_E1_biz_org = 27
	Const S074_E1_deal_type = 28
	Const S074_E1_pay_meth = 29
	Const S074_E1_pay_dur = 30
	Const S074_E1_pay_day = 31
	Const S074_E1_vat_inc_flag = 32
	Const S074_E1_vat_type = 33
	Const S074_E1_vat_rate = 34
	Const S074_E1_trans_meth = 35
	Const S074_E1_trans_lt = 36
	Const S074_E1_sale_amt = 37
	Const S074_E1_capital_amt = 38
	Const S074_E1_emp_cnt = 39
	Const S074_E1_bp_grade = 40
	Const S074_E1_comm_rate = 41
	Const S074_E1_addr2 = 42
	Const S074_E1_addr2_eng = 43
	Const S074_E1_addr3_eng = 44
	Const S074_E1_pay_type = 45
	Const S074_E1_pay_terms_txt = 46
	Const S074_E1_credit_mgmt_flag = 47
	Const S074_E1_credit_grp = 48
	Const S074_E1_vat_calc_type = 49
	Const S074_E1_deposit_flag = 50
	Const S074_E1_bp_group = 51
	Const S074_E1_clearance_id = 52
	Const S074_E1_credit_rot_day = 53
	Const S074_E1_gr_insp_type = 54
	Const S074_E1_bp_alias_nm = 55
	Const S074_E1_to_org = 56
	Const S074_E1_to_grp = 57
	Const S074_E1_pay_month = 58
	Const S074_E1_expiry_dt = 59
	Const S074_E1_pur_grp = 60
	Const S074_E1_pur_org = 61
	Const S074_E1_charge_lay_flag = 62
	Const S074_E1_remark1 = 63
	Const S074_E1_remark2 = 64
	Const S074_E1_remark3 = 65
	Const S074_E1_close_day1 = 66
	Const S074_E1_close_day2 = 67
	Const S074_E1_close_day3 = 68
	Const S074_E1_tax_biz_area = 69
	Const S074_E1_cash_rate = 70
	Const S074_E1_pay_type_out = 71
	Const S074_E1_par_bank_cd1_bp = 72
	Const S074_E1_bank_acct_no1_bp = 73
	Const S074_E1_bank_cd1_bp = 74
	Const S074_E1_par_bank_cd2_bp = 75
	Const S074_E1_bank_cd2_bp = 76
	Const S074_E1_bank_acct_no2_bp = 77
	Const S074_E1_par_bank_cd3_bp = 78
	Const S074_E1_bank_cd3_bp = 79
	Const S074_E1_bank_acct_no3_bp = 80
	Const S074_E1_par_bank_cd1 = 81
	Const S074_E1_bank_cd1 = 82
	Const S074_E1_bank_acct_no1 = 83
	Const S074_E1_par_bank_cd2 = 84
	Const S074_E1_bank_cd2 = 85
	Const S074_E1_bank_acct_no2 = 86
	Const S074_E1_par_bank_cd3 = 87
	Const S074_E1_bank_cd3 = 88
	Const S074_E1_bank_acct_no3 = 89
	Const S074_E1_pay_month2 = 90
	Const S074_E1_pay_day2 = 91
	Const S074_E1_pay_month3 = 92
	Const S074_E1_pay_day3 = 93
	Const S074_E1_close_day1_sales = 94
	Const S074_E1_pay_month1_sales = 95
	Const S074_E1_pay_day1_sales = 96
	Const S074_E1_close_day2_sales = 97
	Const S074_E1_pay_month2_sales = 98
	Const S074_E1_pay_day2_sales = 99
	Const S074_E1_close_day3_sales = 100
	Const S074_E1_pay_month3_sales = 101
	Const S074_E1_pay_day3_sales = 102
	Const S074_E1_ext1_qty = 103
	Const S074_E1_ext2_qty = 104
	Const S074_E1_ext3_qty = 105
	Const S074_E1_ext1_amt = 106
	Const S074_E1_ext2_amt = 107
	Const S074_E1_ext3_amt = 108
	Const S074_E1_ext1_cd = 109
	Const S074_E1_ext2_cd = 110
	Const S074_E1_ext3_cd = 111

	' 컬럼추가 사내외구분 으로인해 배열순서 고침 
	Const S074_E1__in_out = 112                           '[사내외구분]
	' (PB5CS41)dll[12/24 ]변경에 따른 배열고침 
	Const S074_E1_ind_type_nm = 120                           '[업종명]
	Const S074_E1_ind_class_nm = 121                          '[업태명]


	Const S074_E1_bp_group_nm = 122                           '[거래처분류명]
	Const S074_E1_b_country = 123                             '[국가명]
	Const S074_E1_b_province_nm = 124                         '[지방명]
	Const S074_E1_trans_meth_nm = 125                         '[운송방법명]
	Const S074_E1_deal_type_nm = 126						   '[판매유형명]
	Const S074_E1_bp_grade_nm = 127                           '[업체평가등급명]
	Const S074_E1_s_credit_limit = 128                        '[여신관리그룹명]
	Const S074_E1_b_sales_grp_nm = 129                        '[영업그룹명]
	Const S074_E1_b_to_grp_nm = 130                           '[수금그룹명]
	Const S074_E1_b_pur_grp_nm = 131                          '[구매그룹명]
	Const S074_E1_vat_type_nm = 132                           '[부가세유형명]
	Const S074_E1_pay_meth_nm = 133                           '[결재방법명]
	Const S074_E1_pay_type_nm = 134                           '[입출금유형명]
	Const S074_E1_tax_area_nm = 135                           '[세금신고사업장명]
	Const S074_E1_b_zip_code = 136                            '[--우편번호]
	Const S074_E1_b_pur_org = 137                             '[--구매조직코드]
	Const S074_E1_b_pur_org_nm = 138                          '[--구매조직명]
	Const S074_E1_vat_inc_flag_nm = 139                       '[--부과세구분명]

	Dim iCommandSent

	Dim PB5CS41 

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status 

	I1_b_biz_partner = Trim(Request("txtBp_cd"))

	Set PB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")

	If CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	End If

	Call PB5CS41.B_LOOKUP_BIZ_PARTNER(gStrGlobalCollection, I1_b_biz_partner ,E1_b_biz_partner)

	If CheckSYSTEMError(Err,True) = True Then 
		Set PB5CS41 = Nothing 
		Exit Sub
	End If   


	Response.Write "<Script language=vbs> " & vbCr           
	Response.Write " Parent.frm1.txtBp_cd.value        = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_cd))    & """" & vbCr       
	Response.Write " Parent.frm1.txtBp_nm.value        = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_nm))    & """" & vbCr        

	Response.Write " Parent.frm1.txtBp_Rgst_No.value   = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_rgst_no)) & """" & vbCr       
	Response.Write " Parent.frm1.txtCust_full_nm.value = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_full_nm)) & """" & vbCr        
	Response.Write " Parent.frm1.txtCust_nm.value      = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_nm))      & """" & vbCr 
	Response.Write " Parent.frm1.txtCust_eng_nm.value  = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_eng_nm))  & """" & vbCr       
	Response.Write " Parent.frm1.txtRepre_nm.value     = """ & ConvSPChars(E1_b_biz_partner(S074_E1_repre_nm))   & """" & vbCr        
	Response.Write " Parent.frm1.txtRepre_Rgst.value   = """ & ConvSPChars(E1_b_biz_partner(S074_E1_repre_rgst_no)) & """" & vbCr 

	Response.Write " Parent.frm1.txtInd_Type.value     = """ & ConvSPChars(E1_b_biz_partner(S074_E1_ind_type))      & """" & vbCr       
	Response.Write " Parent.frm1.txtInd_TypeNm.value   = """ & ConvSPChars(E1_b_biz_partner(S074_E1_ind_type_nm))   & """" & vbCr            

	Response.Write " Parent.frm1.txtInd_Class.value    = """ & ConvSPChars(E1_b_biz_partner(S074_E1_ind_class))      & """" & vbCr     
	Response.Write " Parent.frm1.txtInd_ClassNm.value  = """ & ConvSPChars(E1_b_biz_partner(S074_E1_ind_class_nm))  & """" & vbCr               

	Response.Write " Parent.frm1.txtZIP_cd.value			= """ & ConvSPChars(E1_b_biz_partner(S074_E1_zip_cd))		& """" & vbCr     
	Response.Write " Parent.frm1.txtADDR1.value			= """ & ConvSPChars(E1_b_biz_partner(S074_E1_addr1))		& """" & vbCr     
	Response.Write " Parent.frm1.txtADDR2.value			= """ & ConvSPChars(E1_b_biz_partner(S074_E1_addr2))		& """" & vbCr     
	Response.Write " Parent.frm1.txtADDR1_Eng.value		= """ & ConvSPChars(E1_b_biz_partner(S074_E1_addr1_eng))    & """" & vbCr
	Response.Write " Parent.frm1.txtADDR2_Eng.value		= """ & ConvSPChars(E1_b_biz_partner(S074_E1_addr2_eng))    & """" & vbCr     
	Response.Write " Parent.frm1.txtADDR3_Eng.value		= """ & ConvSPChars(E1_b_biz_partner(S074_E1_addr3_eng))    & """" & vbCr                  

	Response.Write " Parent.frm1.txtChgCause.value   = """ & ""   & """" & vbCr            

    Response.Write "</Script> "                        & vbCr      

	set PB5CS41 = Nothing 

End Sub

%>

