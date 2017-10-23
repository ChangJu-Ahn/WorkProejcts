<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S7111MB1																	*
'*  4. Program Name         : NEGO 등록																	*
'*  5. Program Desc         : NEGO Query Transaction 처리용 ASP											*
'*  6. Comproxy List        : PSAG111.dll, PSAG119.dll               									*
'*  7. Modified date(First) : 2000/05/09																*
'*  8. Modified date(Last)  : 2000/05/09																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Seo jin kyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/05/09 : 화면 design												*
'*                            2. 2002/06/26 : 

'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%        
    On Error Resume Next                                                             '☜: Protect system from crashing 

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")    
	Call HideStatusWnd                                                                 '☜: Hide Processing message
    
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query			
             Call SubBizQuery()
        Case CStr(UID_M0002)													     '☜: Save
			 Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜:  Delete
             Call SubBizDelete()
    End Select

Sub SubBizQuery()

	Dim pvCommand 
	Dim I1_s_nego_nego_no 
	Dim EG1_E1_exp_grp 

	Const S708_EG1_nego_no = 0
	Const S708_EG1_nego_doc_no = 1
	Const S708_EG1_nego_dt = 2
	Const S708_EG1_pay_expiry_dt = 3
	Const S708_EG1_nego_pub_zone = 4
	Const S708_EG1_manufacturer = 5
	Const S708_EG1_agent = 6
	Const S708_EG1_cur = 7
	Const S708_EG1_nego_doc_amt = 8
	Const S708_EG1_nego_amt_txt = 9
	Const S708_EG1_xch_rate = 10
	Const S708_EG1_xch_comn_rate = 11
	Const S708_EG1_nego_loc_amt = 12
	Const S708_EG1_incoterms = 13
	Const S708_EG1_pay_meth = 14
	Const S708_EG1_pay_dur = 15
	Const S708_EG1_nego_req_dt = 16
	Const S708_EG1_flaw_exist = 17
	Const S708_EG1_pay_dt = 18
	Const S708_EG1_pay_type = 19
	Const S708_EG1_bl_doc_no = 20
	Const S708_EG1_bas_doc_amt = 21
	Const S708_EG1_lc_doc_no = 22
	Const S708_EG1_lc_amend_seq = 23
	Const S708_EG1_lc_open_dt = 24
	Const S708_EG1_lc_expiry_dt = 25
	Const S708_EG1_latest_ship_dt = 26
	Const S708_EG1_nego_type = 27
	Const S708_EG1_adv_no = 28
	Const S708_EG1_remarks1 = 29
	Const S708_EG1_remarks2 = 30
	Const S708_EG1_ext1_qty = 31
	Const S708_EG1_ext1_amt = 32
	Const S708_EG1_bill_no = 33
	Const S708_EG1_posting_flg = 34
	Const S708_EG1_cost_cd = 35
	Const S708_EG1_biz_area = 36
	Const S708_EG1_ext2_qty = 37
	Const S708_EG1_ext3_qty = 38
	Const S708_EG1_ext2_amt = 39
	Const S708_EG1_ext3_amt = 40
	Const S708_EG1_ext1_cd = 41
	Const S708_EG1_ext2_cd = 42
	Const S708_EG1_ext3_cd = 43
	Const S708_EG1_xch_rate_op = 44
	Const S708_EG1_bank_cd = 45
	Const S708_EG1_bank_acct_no = 46
	Const S708_EG1_bp_cd_appli = 47
	Const S708_EG1_bp_nm_appli = 48
	Const S708_EG1_bp_cd_benff = 49
	Const S708_EG1_bp_nm_benff = 50
	Const S708_EG1_sales_grp = 51

	Const S708_EG1_sales_grp_nm = 52
	Const S708_EG1_sales_org = 53
	Const S708_EG1_sales_org_nm = 54
	Const S708_EG1_lc_no = 55
	Const S708_EG1_bank_cd_cd = 56
	Const S708_EG1_bank_nm = 57
	Const S708_EG1_bank_cd_open = 58
	Const S708_EG1_bank_nm_open = 59

	Const S708_EG1_minor_nm_pay_type = 60
	Const S708_EG1_minor_nm_nego_type = 61
	Const S708_EG1_minor_nm_incoterms = 62
	Const S708_EG1_minor_nm_pay_meth = 63
	Const S708_EG1_bp_nm_manufacturer = 64
	Const S708_EG1_bp_nm_agent = 65
	Const S708_EG1_bank_nm_ngb = 66

	
	Dim PSAG119
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Request("txtPrevNext")
		Case "PREV"
			pvCommand = "PREV"
		Case "NEXT"
			pvCommand = "NEXT"
		Case Else 
			pvCommand = "LOOKUP"
	End Select	
	
		
	I1_s_nego_nego_no = Trim(Request("txtNEGONo"))	
	
    Set PSAG119 = Server.CreateObject("PSAG119.cSLkNegoSvr")
	
	if CheckSYSTEMError(Err,True) = True Then 
        Response.Write "<Script language=vbs>  " & vbCr   
        Response.Write "   Parent.frm1.txtNegoNo.focus " & vbCr    
        Response.Write "</Script>      " & vbCr	
		Exit Sub
	end if
	
	Call PSAG119.S_LOOKUP_NEGO_SVR (gStrGlobalCollection, pvCommand,I1_s_nego_nego_no, EG1_E1_exp_grp)	
    
    If CheckSYSTEMError(Err,True) = True Then 		
		Set PSAG119 = Nothing
        Response.Write "<Script language=vbs>  " & vbCr   
        Response.Write "   Parent.frm1.txtNegoNo.focus " & vbCr    
        Response.Write "</Script>      " & vbCr		
		Exit Sub		
	end if	
	
	Set PSAG119 = Nothing
		
	 
	Response.Write "<Script language=vbs> " & vbCr    
	   
    Response.Write " Parent.frm1.txtCurrency.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_cur))                 & """" & vbCr    
    Response.Write " parent.CurFormatNumericOCX  "																	   & vbCr   
    Response.Write " Parent.frm1.txtNegoNo.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_no))             & """" & vbCr  
    Response.Write " Parent.frm1.txtHNEGONo.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_no))             & """" & vbCr  
    
    
    'Tab 1 : NEGO 정보			  
    Response.Write " Parent.frm1.txtNegoNo1.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_no))             & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoDocNo.value  = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_doc_no))         & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoType.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_type ))          & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoTypeNm.value = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_minor_nm_nego_type))  & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoBank.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_cd_cd))          & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoBankNm.value = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_nm))             & """" & vbCr    
    Response.Write " Parent.frm1.txtNegoDt.text     = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_nego_dt))     & """" & vbCr    	
    
    Response.Write " Parent.frm1.rdoPostingflg1.disabled = false " & vbCr    	
    Response.Write " Parent.frm1.rdoPostingflg2.disabled = false " & vbCr    	

	Response.Write " Parent.frm1.txtNegoDocAmt.text     = """ & UNINumClientFormat(EG1_E1_exp_grp(S708_EG1_nego_doc_amt), ggAmtOfMoney.DecPoint, 0)      & """" & vbCr    	
	Response.Write " Parent.frm1.txtNegoAmtTxt.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_amt_txt))                                       & """" & vbCr    
	Response.Write " Parent.frm1.txtXchRate.Text         = """ & UNINumClientFormat(EG1_E1_exp_grp(S708_EG1_xch_rate), ggExchRate.DecPoint, 0)            & """" & vbCr    	
	
	Response.Write " Parent.frm1.txtNegoLocAmt.Text      = """ & UniConvNumberDBToCompany(EG1_E1_exp_grp(S708_EG1_nego_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)  & """" & vbCr    
	
	
	Response.Write " Parent.frm1.txtSalesGroup.value          = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_sales_grp))                    & """" & vbCr    
	Response.Write " Parent.frm1.txtSalesGroupNm.value          = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_sales_grp_nm))               & """" & vbCr    
	Response.Write " Parent.frm1.txtPayExpiryDt.text          = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_pay_expiry_dt))       & """" & vbCr    
	Response.Write " Parent.frm1.txtNegoReqDt.text          = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_nego_req_dt ))          & """" & vbCr    
	
	If EG1_E1_exp_grp(S708_EG1_flaw_exist )  = "Y" Then
		Response.Write " Parent.frm1.rdoFlawExist1.Checked  = True " & vbCr   					 		
	ElseIf EG1_E1_exp_grp(S708_EG1_flaw_exist )  = "N" Then	
		Response.Write " Parent.frm1.rdoPostingflg2.Checked = True " & vbCr    		
	End If


	Response.Write " Parent.frm1.txtCollectType.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_pay_type))            & """" & vbCr    
	Response.Write " Parent.frm1.txtCollectTypeNm.value  = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_minor_nm_pay_type))   & """" & vbCr    
	Response.Write " Parent.frm1.txtPayDt.text          = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_pay_dt ))     & """" & vbCr    
	

	Response.Write " Parent.frm1.txtAccountNo.value      = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_acct_no))        & """" & vbCr    
	Response.Write " Parent.frm1.txtIncomeBank.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_cd))             & """" & vbCr    
	Response.Write " Parent.frm1.txtIncomeBankNm.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_nm_ngb ))        & """" & vbCr    
	Response.Write " Parent.frm1.txtNegoPubZone.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_nego_pub_zone))       & """" & vbCr    
	Response.Write " Parent.frm1.txtRemarks1.value       = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_remarks1))            & """" & vbCr    
	Response.Write " Parent.frm1.txtRemarks2.value       = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_remarks2))            & """" & vbCr    
	

		'Tab 2 
	Response.Write " Parent.frm1.txtXchCommRate.text    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_xch_comn_rate))       & """" & vbCr    
	Response.Write " Parent.frm1.txtAdvNo.value          = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_adv_no))              & """" & vbCr    
	Response.Write " Parent.frm1.txtBillNo.value         = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bill_no))             & """" & vbCr    
	
	Response.Write " Parent.frm1.txtBLNo.value           = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bl_doc_no))           & """" & vbCr    
	Response.Write " Parent.frm1.txtLCNo.value           = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_lc_no))               & """" & vbCr    
	Response.Write " Parent.frm1.txtLCDocNo.value        = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_lc_doc_no))           & """" & vbCr    
	Response.Write " Parent.frm1.txtLCAmendSeq.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_lc_amend_seq))        & """" & vbCr    
	Response.Write " Parent.frm1.txtBaseCurrency.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_cur))                 & """" & vbCr    
	Response.Write " Parent.frm1.txtBaseDocAmt.text     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_remarks2))            & """" & vbCr    
	
	
	Response.Write " Parent.frm1.txtBaseDocAmt.text      = """ & UNINumClientFormat(EG1_E1_exp_grp(S708_EG1_bas_doc_amt), ggAmtOfMoney.DecPoint, 0)      & """" & vbCr    	
	Response.Write " Parent.frm1.txtLatestShipDt.text    = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_latest_ship_dt ))                           & """" & vbCr    
	Response.Write " Parent.frm1.txtOpenDt.text          = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_lc_open_dt ))                               & """" & vbCr    
	Response.Write " Parent.frm1.txtExpireDt.text        = """ & UNIDateClientFormat(EG1_E1_exp_grp(S708_EG1_lc_expiry_dt ))                             & """" & vbCr    
	
	Response.Write " Parent.frm1.txtOpenBank.value       = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_cd_open))         & """" & vbCr    
	Response.Write " Parent.frm1.txtOpenBankNm.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bank_nm_open))         & """" & vbCr    
	Response.Write " Parent.frm1.txtIncoterms.value      = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_incoterms))            & """" & vbCr    
	Response.Write " Parent.frm1.txtPayTerms.value       = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_pay_meth))             & """" & vbCr    
	Response.Write " Parent.frm1.txtPayTermsNm.value     = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_minor_nm_pay_meth))    & """" & vbCr    	
	

	
	Response.Write " Parent.frm1.txtPayDur.value         = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_pay_dur))              & """" & vbCr    
	Response.Write " Parent.frm1.txtApplicant.value      = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_cd_appli))          & """" & vbCr    	
	Response.Write " Parent.frm1.txtApplicantNm.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_nm_appli))          & """" & vbCr  
		
	Response.Write " Parent.frm1.txtBeneficiary.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_cd_benff))          & """" & vbCr    
	Response.Write " Parent.frm1.txtBeneficiaryNm.value  = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_nm_benff))          & """" & vbCr    
	Response.Write " Parent.frm1.txtAgent.value          = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_agent))                & """" & vbCr    
	Response.Write " Parent.frm1.txtAgentNm.value        = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_nm_agent))          & """" & vbCr    
	Response.Write " Parent.frm1.txtManufacturer.value   = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_manufacturer))         & """" & vbCr    
	Response.Write " Parent.frm1.txtManufacturerNm.value = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_bp_nm_manufacturer))   & """" & vbCr    

	If Trim(EG1_E1_exp_grp(S708_EG1_posting_flg))  = "Y" Then
		Response.Write " Parent.frm1.rdoPostingflg1.Checked  =  True "                                                   & vbCr 
		Response.Write " Parent.frm1.btnPosting.value        = ""확정취소"""			    		   		     & vbCr		
'		Response.Write " Parent.ProtectBody  "                                                   & vbCr     
	ElseIf Trim(EG1_E1_exp_grp(S708_EG1_posting_flg))  = "N" Then	
		Response.Write " Parent.frm1.rdoPostingflg2.Checked  =  True "												     & vbCr    
		Response.Write " Parent.frm1.btnPosting.value        = ""확정"""										 & vbCr		
'		Response.Write " Parent.ReleaseBody  "
	End If


    Response.Write " Parent.DbQueryOk "										 & vbCr  
    
    If EG1_E1_exp_grp(S708_EG1_flaw_exist)  = "Y" Then 
		Response.Write " parent.rdoFlawExist1_OnClick " & vbCr    				
		Response.Write " Parent.frm1.rdoFlawExist1.Checked = True " & vbCr   					 				
		
	ElseIf EG1_E1_exp_grp(S708_EG1_flaw_exist)  = "N" Then		
		Response.Write " parent.rdoFlawExist2_OnClick " & vbCr    		
		Response.Write " Parent.frm1.rdoFlawExist2.Checked = True " & vbCr  				
	End If				
		
	Response.Write " Parent.frm1.txtExchRateOp.value    = """ & ConvSPChars(EG1_E1_exp_grp(S708_EG1_xch_rate_op))     & """" & vbCr    
	Response.Write " Parent.frm1.txtHNEGONo.value       = """ & ConvSPChars(Request("txtNEGONo"))                     & """" & vbCr    
    Response.Write "</Script> "																			              & vbCr
	

End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save DB data
'============================================================================================================
Sub SubBizSave()
	
	Dim lgIntFlgMode
	Dim pvCommand
	Dim I1_s_nego
	Const S700_I1_nego_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_nego
	Const S700_I1_nego_doc_no = 1
	Const S700_I1_nego_dt = 2
	Const S700_I1_pay_expiry_dt = 3
	Const S700_I1_nego_pub_zone = 4
	Const S700_I1_cur = 5
	Const S700_I1_nego_doc_amt = 6
	Const S700_I1_nego_amt_txt = 7
	Const S700_I1_xch_rate = 8
	Const S700_I1_xch_comn_rate = 9
	Const S700_I1_nego_loc_amt = 10
	Const S700_I1_nego_req_dt = 11
	Const S700_I1_flaw_exist = 12
	Const S700_I1_pay_dt = 13
	Const S700_I1_pay_type = 14
	Const S700_I1_nego_type = 15
	Const S700_I1_adv_no = 16
	Const S700_I1_remarks1 = 17
	Const S700_I1_remarks2 = 18
	Const S700_I1_bas_doc_amt = 19
	Const S700_I1_ext1_amt = 20
	Const S700_I1_ext2_amt = 21
	Const S700_I1_ext3_amt = 22
	Const S700_I1_ext1_cd = 23
	Const S700_I1_ext2_cd = 24
	Const S700_I1_ext3_cd = 25
	Const S700_I1_ext1_qty = 26
	Const S700_I1_ext2_qty = 27
	Const S700_I1_ext3_qty = 28
	
	Dim I2_b_sales_grp_sales_grp
	Dim I3_s_lc_hdr_lc_no
	Dim I4_b_bank_bank_cd
	Dim I5_s_bill_hdr_bill_no
	Dim I6_b_bank_acct_acct_no
	Dim I7_b_bank_bank_cd
	Dim E1_s_nego_nego_no
	Dim PSAG111

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Redim I1_s_nego(28)
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 
    
    I1_s_nego(S700_I1_nego_no) = UCase(Trim(Request("txtNEGONo1")))
	I1_s_nego(S700_I1_nego_doc_no) = Trim(Request("txtNegoDocNo"))
	I1_s_nego(S700_I1_nego_type) = UCase(Trim(Request("txtNEGOType")))
	I4_b_bank_bank_cd = UCase(Trim(Request("txtNegoBank")))
	
	If Len(Trim(Request("txtNegoDt"))) Then
		I1_s_nego(S700_I1_nego_dt) = UNIConvDate(Request("txtNegoDt"))			
	End If
	
	I1_s_nego(S700_I1_cur) = UCase(Trim(Request("txtCurrency")))
	

	If Len(Trim(Request("txtNegoDocAmt"))) Then
		I1_s_nego(S700_I1_nego_doc_amt) = UNIConvNum(Request("txtNegoDocAmt"),0)
	End If

	I1_s_nego(S700_I1_nego_amt_txt) = UCase(Trim(Request("txtNegoAmtTxt")))
	I1_s_nego(S700_I1_xch_rate) = UNIConvNum(Request("txtXchRate"),0)
		
	If Len(Trim(Request("txtNegoLocAmt"))) Then
		I1_s_nego(S700_I1_nego_loc_amt) = UNIConvNum(Request("txtNegoLocAmt"),0)
	End If
			
	I2_b_sales_grp_sales_grp = UCase(Trim(Request("txtSalesGroup")))		
	I1_s_nego(S700_I1_pay_expiry_dt) = UNIConvDate(Request("txtPayExpiryDt"))		
	I1_s_nego(S700_I1_nego_req_dt) = UNIConvDate(Request("txtNegoReqDt"))			
	I1_s_nego(S700_I1_flaw_exist) = Request("rdoFlawExist")
	I1_s_nego(S700_I1_pay_type) = UCase(Trim(Request("txtCollectType")))
	I1_s_nego(S700_I1_pay_dt) = UNIConvDate(Request("txtPayDt"))
	
		
	I6_b_bank_acct_acct_no = Trim(Request("txtAccountNo"))
	I7_b_bank_bank_cd = UCase(Trim(Request("txtIncomeBank")))
	
	I1_s_nego(S700_I1_nego_pub_zone) = Trim(Request("txtNegoPubZone"))
	I1_s_nego(S700_I1_remarks1) = Trim(Request("txtRemarks1"))
	I1_s_nego(S700_I1_remarks2) = Trim(Request("txtRemarks2"))
		


	'Tab 2
		
	If Len(Trim(Request("txtXchCommRate"))) Then
		I1_s_nego(S700_I1_xch_comn_rate) = UNIConvNum(Request("txtXchCommRate"),0)
	End If
		
	I1_s_nego(S700_I1_adv_no) = Trim(Request("txtAdvNo"))
	I5_s_bill_hdr_bill_no =	UCase(Trim(Request("txtBillNo")))	
	I3_s_lc_hdr_lc_no = UCase(Trim(Request("txtLCNo")))
	I1_s_nego(S700_I1_bas_doc_amt) = UNIConvNum(Request("txtBaseDocAmt"),0)
		
	If lgIntFlgMode = OPMD_CMODE Then
		pvCommand = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		pvCommand = "UPDATE"
	End If

    Set PSAG111 = Server.CreateObject("PSAG111.cSNegoSvr")
	
	If CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	End if
   
	E1_s_nego_nego_no =  PSAG111.S_MAINT_NEGO_SVR(gStrGlobalCollection, pvCommand,I1_s_nego,  I2_b_sales_grp_sales_grp , _
		I3_s_lc_hdr_lc_no, I4_b_bank_bank_cd, I5_s_bill_hdr_bill_no, _
		I6_b_bank_acct_acct_no, I7_b_bank_bank_cd )
    
    If CheckSYSTEMError(Err,True) = True Then 		
		Set PSAG111 = Nothing
		Exit Sub		
	end if	
	
	Set PSAG111 = Nothing 	
	
	Response.Write "<Script language=vbs> " & vbCr    	
	Response.Write " Parent.frm1.txtNEGONo.value    = """ & ConvSPChars(E1_s_nego_nego_no)     & """" & vbCr    
    Response.Write " Parent.DbSaveOk "																			    	    & vbCr   
    Response.Write "</Script> "						

End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

	Dim pvCommand
	Dim I1_s_nego
	Const S700_I1_nego_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_nego
		
	Dim I2_b_sales_grp_sales_grp
	Dim I3_s_lc_hdr_lc_no
	Dim I4_b_bank_bank_cd
	Dim I5_s_bill_hdr_bill_no
	Dim I6_b_bank_acct_acct_no
	Dim I7_b_bank_bank_cd
	
	Dim PSAG111

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Redim I1_s_nego(28)

	I1_s_nego(S700_I1_nego_no) = UCase(Trim(Request("txtNEGONo")))
	
	pvCommand = "DELETE"
	
	Set PSAG111 = Server.CreateObject("PSAG111.cSNegoSvr")	
	If CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	End if
   
	Call PSAG111.S_MAINT_NEGO_SVR(gStrGlobalCollection, pvCommand,I1_s_nego,  I2_b_sales_grp_sales_grp , _
		I3_s_lc_hdr_lc_no, I4_b_bank_bank_cd, I5_s_bill_hdr_bill_no, _
		I6_b_bank_acct_acct_no, I7_b_bank_bank_cd )
		    
    If CheckSYSTEMError(Err,True) = True Then 		
		Set PSAG111 = Nothing
		Exit Sub		
	end if		
	
	Set PSAG111 = Nothing 
	
	Response.Write "<Script language=vbs> "    & vbCr    		
    Response.Write " Parent.DbDeleteOk "       & vbCr   
    Response.Write "</Script> "				
    
End Sub
%>
