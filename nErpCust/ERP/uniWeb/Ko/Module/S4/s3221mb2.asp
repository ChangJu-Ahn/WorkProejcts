<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221mb2.asp																*
'*  4. Program Name         : Local L/C Amend 등록														*
'*  5. Program Desc         : Local L/C Amend 등록														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/24																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/31 : 화면 design												*
'*							  2. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%   
    On Error Resume Next                                                             '☜: Protect system from crashing 
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
	Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
	Call HideStatusWnd          
																					 '☜: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query			
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update, 
             Call SubBizSave()
        Case CStr(UID_M0003)														  '☜:  Delete
             Call SubBizDelete()
    End Select

Sub SubBizQuery()	    

	Dim pvCommand 
	Dim I1_s_lc_amend_hdr 
	Const S375_I1_lc_amd_no = 0    'imp s_lc_amend_hdr
	Const S375_I1_lc_kind = 1

	Dim E1_b_minor 
	Const S375_E1_minor_nm = 0    'exp_pay_meth_nm b_minor
	
	Dim E2_b_minor 
	Const S375_E2_minor_nm = 0    'exp_incoterms_nm b_minor	
	
	Dim E3_b_minor 
	Const S375_E3_minor_nm = 0    'exp_be_transport_nm b_minor	Dim E4_b_minor 
	
	Dim E4_b_minor 
	Const S375_E4_minor_nm = 0    'exp_be_discharge_port_nm b_minor
	
	Dim E5_b_minor 
	Const S375_E5_minor_nm = 0    'exp_be_loading_port_nm b_minor
	
	Dim E6_b_minor 
	Const S375_E6_minor_nm = 0    'exp_at_transport_nm b_minor
	
	Dim E7_b_minor 
	Const S375_E7_minor_nm = 0    'exp_at_discharge_port_nm b_minor
	
	Dim E8_b_minor 
	Const S375_E8_minor_nm = 0    'exp_at_loading_port_nm b_minor
	
	Dim E9_s_lc_hdr 
	Const S375_E9_lc_no = 0    'exp s_lc_hdr
	Const S375_E9_so_no = 1
	Const S375_E9_incoterms = 2
	Const S375_E9_pay_meth = 3

	Dim E10_b_sales_org 
	Const S375_E10_sales_org = 0    'exp b_sales_org
	Const S375_E10_sales_org_nm = 1

	Dim E11_b_sales_grp 	
	Const S375_E11_sales_grp = 0    'exp b_sales_grp
	Const S375_E11_sales_grp_nm = 1

	Dim E12_s_lc_amend_hdr 	
	Const S375_E12_lc_amd_no = 0    'exp s_lc_amend_hdr
	Const S375_E12_lc_doc_no = 1
	Const S375_E12_lc_amend_seq = 2
	Const S375_E12_adv_no = 3
	Const S375_E12_pre_adv_ref = 4
	Const S375_E12_open_dt = 5
	Const S375_E12_be_expiry_dt = 6
	Const S375_E12_at_expiry_dt = 7
	Const S375_E12_manufacturer = 8
	Const S375_E12_agent = 9
	Const S375_E12_amend_dt = 10
	Const S375_E12_amend_req_dt = 11
	Const S375_E12_cur = 12
	Const S375_E12_be_lc_amt = 13
	Const S375_E12_at_lc_amt = 14
	Const S375_E12_at_xch_rate = 15
	Const S375_E12_inc_amt = 16
	Const S375_E12_dec_amt = 17
	Const S375_E12_be_loc_amt = 18
	Const S375_E12_at_loc_amt = 19
	Const S375_E12_be_latest_ship_dt = 20
	Const S375_E12_at_latest_ship_dt = 21
	Const S375_E12_be_xch_rate = 22
	Const S375_E12_remark = 23
	Const S375_E12_be_loading_port = 24
	Const S375_E12_at_loading_port = 25
	Const S375_E12_be_dischge_port = 26
	Const S375_E12_at_dischge_port = 27
	Const S375_E12_be_transport = 28
	Const S375_E12_at_transport = 29
	Const S375_E12_remark2 = 30
	Const S375_E12_be_partial_ship_flag = 31
	Const S375_E12_at_partial_ship_flag = 32
	Const S375_E12_lc_kind = 33
	Const S375_E12_be_trnshp_flag = 34
	Const S375_E12_at_trnshp_flag = 35
	Const S375_E12_be_transfer_flag = 36
	Const S375_E12_at_transfer_flag = 37
	Const S375_E12_advise_bank = 38
	Const S375_E12_ext1_qty = 39
	Const S375_E12_ext2_qty = 40
	Const S375_E12_ext3_qty = 41
	Const S375_E12_ext1_amt = 42
	Const S375_E12_ext2_amt = 43
	Const S375_E12_ext3_amt = 44
	Const S375_E12_ext1_cd = 45
	Const S375_E12_ext2_cd = 46
	Const S375_E12_ext3_cd = 47
	Const S375_E12_xch_rate_op = 48

	Dim E13_b_biz_partner 	
	Const S375_E13_bp_nm = 0    'exp_agent b_biz_partner

	Dim E14_b_biz_partner 	    
	Const S375_E14_bp_nm = 0    'exp_manufacturer b_biz_partner

	Dim E15_b_biz_partner 	    
	Const S375_E15_bp_nm = 0    'exp_applicant b_biz_partner
	Const S375_E15_bp_cd = 1

	Dim E16_b_bank 	    
	Const S375_E16_bank_nm = 0    'exp_issue b_bank
	Const S375_E16_bank_cd = 1

	Dim E17_b_bank 
	Const S375_E17_bank_nm = 0    'exp_advise b_bank

	Dim E18_b_biz_partner 	    
	Const S375_E18_bp_nm = 0    'exp_beneficiary b_biz_partner
	Const S375_E18_bp_cd = 1
	
	Dim E19_s_lc_dtl 
	Const S375_E19_lc_amt = 0    'exp_tot s_lc_dtl
	Const S375_E19_lc_loc_amt = 1

	Dim I1_currency 
	Dim I2_currency 
	Dim I3_apprl_dt 
	


	Dim E1_b_daily_exchange_rate 
	Const B253_E1_std_rate = 0	
	Const B253_E1_multi_divide = 1    

	Dim PS4G219
	Dim PB0C004
	
	On Error Resume Next
    Err.Clear
    
    If Request("txtLCAmdNo") = "" Then										
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Response.Write "<Script language=vbs> " & vbCr   
        Response.Write "   Parent.frm1.txtLCAmdNo.focus " & vbCr    
		Response.Write "</Script> " & vbCr      
		Exit Sub
	End If
	

	Select Case Request("txtPrevNext")
		Case "PREV"
			pvCommand = "PREV"
		Case "NEXT"
			pvCommand = "NEXT"
		Case Else 
			pvCommand = "QUERY"
	End Select	
	
	Redim I1_s_lc_amend_hdr(1)
	
	
	I1_s_lc_amend_hdr(S375_I1_lc_amd_no) = Trim(Request("txtLCAmdNo"))
	I1_s_lc_amend_hdr(S375_I1_lc_kind) = "L"
	
	
	Set PS4G219 = Server.CreateObject("PS4G219.cSLkLcAmendHdrSvr")			
	if CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	end If
	
	
	call PS4G219.S_LOOKUP_LC_AMEND_HDR_SVR(gStrGlobalCollection, pvCommand , I1_s_lc_amend_hdr , _
			E1_b_minor,			E2_b_minor,			E3_b_minor,			E4_b_minor, _
			E5_b_minor,			E6_b_minor,			E7_b_minor,			E8_b_minor, _
			E9_s_lc_hdr,		E10_b_sales_org,	E11_b_sales_grp,	E12_s_lc_amend_hdr,  _
			E13_b_biz_partner,	E14_b_biz_partner,	E15_b_biz_partner,	E16_b_bank,  _
			E17_b_bank,			E18_b_biz_partner,	E19_s_lc_dtl )
						
	if CheckSYSTEMError(Err,True) = True Then 
		Set PS4G219 = nothing
		Response.Write "<Script language=vbs> " & vbCr   
        Response.Write "   Parent.frm1.txtLCAmdNo.focus " & vbCr    
		Response.Write "</Script> " & vbCr       
		Exit Sub
	end If	

	Response.Write "<Script language=vbs> " & vbCr   
	'##### Rounding Logic #####    
    Response.Write " Parent.frm1.txtAtCurrency1.value		= """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur)) & """" & vbCr 
    Response.Write " Parent.frm1.txtAtCurrency2.value        = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur)) & """" & vbCr       
    Response.Write " Parent.CurFormatNumericOCX "									                                     & vbCr       
	'##########################    		
		
    Response.Write " Parent.frm1.txtLCAmdNo.value           = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amd_no))						 & """" & vbCr        
    Response.Write " Parent.frm1.txtLCAmdNo1.value		    = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amd_no))						 & """" & vbCr           
    Response.Write " Parent.frm1.txtLCNo.value			    = """ & ConvSPChars(E9_s_lc_hdr(S375_E9_lc_no))									 & """" & vbCr            
    Response.Write " Parent.frm1.txtLCDocNo.value		    = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_doc_no))						 & """" & vbCr 
    Response.Write " Parent.frm1.txtLCAmendSeq.value	    = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amend_seq))					 & """" & vbCr               
    
    Response.Write " Parent.frm1.txtAmendDt.text		    = """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_amend_dt)))  & """" & vbCr             
    
    Response.Write " Parent.frm1.txtApplicant.value         = """ & ConvSPChars(E15_b_biz_partner(S375_E15_bp_cd))							 & """" & vbCr        
    Response.Write " Parent.frm1.txtApplicantNm.value       = """ & ConvSPChars(E15_b_biz_partner(S375_E15_bp_nm))							 & """" & vbCr     
    Response.Write " Parent.frm1.txtBeneficiary.value		= """ & ConvSPChars(E18_b_biz_partner(S375_E18_bp_cd))							 & """" & vbCr       
    Response.Write " Parent.frm1.txtBeneficiaryNm.value     = """ & ConvSPChars(E18_b_biz_partner(S375_E18_bp_nm))							 & """" & vbCr    
    
    Response.Write " Parent.frm1.txtAtCurrency1.value       = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur))							 & """" & vbCr    
    Response.Write " Parent.frm1.txtAtCurrency2.value       = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur))							 & """" & vbCr    
    
        
	If Trim(E12_s_lc_amend_hdr(S375_E12_inc_amt)) > "0" then
		Response.Write " Parent.frm1.rdoAtDocAmt1.Checked = True "							             & vbCr    
		Response.Write " Parent.frm1.txtAmendAmt.value = """ & UNINumClientFormatByCurrency(E12_s_lc_amend_hdr(S375_E12_inc_amt), E12_s_lc_amend_hdr(S375_E12_cur),ggAmtOfMoneyNo ) & """" & vbCr 
	ElseIf Trim(E12_s_lc_amend_hdr(S375_E12_dec_amt)) > "0" then
		Response.Write " Parent.frm1.rdoAtDocAmt2.Checked = True "							             & vbCr    
		Response.Write " Parent.frm1.txtAmendAmt.value = """ & UNINumClientFormatByCurrency(E12_s_lc_amend_hdr(S375_E12_dec_amt), E12_s_lc_amend_hdr(S375_E12_cur),ggAmtOfMoneyNo ) & """" & vbCr 	
	Else
		Response.Write " Parent.frm1.txtAmendAmt.value = """ & UNINumClientFormatByCurrency(0, E12_s_lc_amend_hdr(S375_E12_cur),ggAmtOfMoneyNo ) & """" & vbCr 	
	End If  
    
	'##### Rounding Logic #####
    Response.Write " Parent.frm1.txtAtDocAmt.value    = """ & UNINumClientFormatByCurrency(E12_s_lc_amend_hdr(S375_E12_at_lc_amt), E12_s_lc_amend_hdr(S375_E12_cur), ggAmtOfMoneyNo)							 & """" & vbCr        
    Response.Write " Parent.frm1.txtBeDocAmt.value	  = """ & UNINumClientFormatByCurrency(E12_s_lc_amend_hdr(S375_E12_be_lc_amt), E12_s_lc_amend_hdr(S375_E12_cur), ggAmtOfMoneyNo)							 & """" & vbCr           
	'##########################		    
        	
    Response.Write " Parent.frm1.txtAtXchRate.value     = """ & UNINumClientFormat(E12_s_lc_amend_hdr(S375_E12_at_xch_rate), ggExchRate.DecPoint, 0)  & """" & vbCr 
    Response.Write " Parent.frm1.txtBeXchRate.value		= """ & UNINumClientFormat(E12_s_lc_amend_hdr(S375_E12_be_xch_rate), ggExchRate.DecPoint, 0)  & """" & vbCr       
    
    '##### Rounding Logic #####
    Response.Write " Parent.frm1.txtAtLocAmt.value    = """ & UniConvNumberDBToCompany(E12_s_lc_amend_hdr(S375_E12_at_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)							 & """" & vbCr        
    Response.Write " Parent.frm1.txtBeLocAmt.value	  = """ & UniConvNumberDBToCompany(E12_s_lc_amend_hdr(S375_E12_be_loc_amt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)							 & """" & vbCr           
	'##########################		
    
    Response.Write " Parent.frm1.txtAtExpireDt.text		= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_at_expiry_dt)))  & """" & vbCr            
    Response.Write " Parent.frm1.txtHExpiryDt.value		= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_at_expiry_dt)))  & """" & vbCr            
    
    Response.Write " Parent.frm1.txtBeExpireDt.text		= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_be_expiry_dt)))  & """" & vbCr                    
    Response.Write " Parent.frm1.txtAtLatestShipDt.text	= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_at_latest_ship_dt)))  & """" & vbCr                
    Response.Write " Parent.frm1.txtHLatestShipDt.value	= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_at_latest_ship_dt)))  & """" & vbCr                
        
    Response.Write " Parent.frm1.txtBeLatestShipDt.text	= """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_be_latest_ship_dt)))  & """" & vbCr                
    

	If E12_s_lc_amend_hdr(S375_E12_at_partial_ship_flag) = "Y" Then        
		Response.Write " Parent.frm1.rdoAtPartialShip1.Checked = True "							             & vbCr    
    ElseIf E12_s_lc_amend_hdr(S375_E12_at_partial_ship_flag) = "N" Then        
		Response.Write " Parent.frm1.rdoAtPartialShip2.Checked = True "							             & vbCr    		
    End If    
		
   	Response.Write " Parent.frm1.txtBePartialShip.value  = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_be_partial_ship_flag))			 & """" & vbCr 		
	Response.Write " Parent.frm1.txtOpenDt.text			 = """ & UNIDateClientFormat(ConvSPChars(E12_s_lc_amend_hdr(S375_E12_open_dt)))  & """" & vbCr                
	
	Response.Write " Parent.frm1.txtAdvNo.value			 = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_adv_no))				& """" & vbCr 	
	Response.Write " Parent.frm1.txtRef.value			 = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_remark))				& """" & vbCr 	
	Response.Write " Parent.frm1.txtAdvBank.value		 = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_advise_bank))			& """" & vbCr 	
    Response.Write " Parent.frm1.txtAdvBankNm.value		 = """ & ConvSPChars(E17_b_bank(S375_E17_bank_nm))						& """" & vbCr 	
    Response.Write " Parent.frm1.txtOpenBank.value		 = """ & ConvSPChars(E16_b_bank(S375_E16_bank_cd))						& """" & vbCr 	
    Response.Write " Parent.frm1.txtOpenBankNm.value     = """ & ConvSPChars(E16_b_bank(S375_E16_bank_nm))						& """" & vbCr 	        
    Response.Write " Parent.frm1.txtSalesGroup.value     = """ & ConvSPChars(E11_b_sales_grp(S375_E11_sales_grp))				& """" & vbCr 	                
    Response.Write " Parent.frm1.txtSalesGroupNm.value   = """ & ConvSPChars(E11_b_sales_grp(S375_E11_sales_grp_nm))			& """" & vbCr 	        
    
    Response.Write " Parent.frm1.txtPreAdvRef.value		 = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_pre_adv_ref))				& """" & vbCr 	        
    
    Response.Write " Parent.frm1.txtHLCAmdNo.value       = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amd_no))		   	    & """" & vbCr 	            
    Response.Write " Parent.frm1.txtExchRateOp.value	 = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_xch_rate_op))														& """" & vbCr 	        
    
    	
    Response.Write " Parent.DbQueryOk "																							    	& vbCr   
    Response.Write " Parent.ProtectXchRate "												                                   			& vbCr   
    Response.Write " Parent.frm1.txtHLCAmdNo.value     = """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amd_no))		   			    & """" & vbCr 	            
    
    Response.Write "</Script> "																											& vbCr          
	Set PS4G219 = Nothing	    


End Sub 

Sub SubBizSave

	Dim iCommandSent
	Dim lgIntFlgMode
	Dim PS4G211
	
	Dim I2_s_lc_amend_hdr
	Dim I3_s_lc_hdr_lc_no
	Dim E2_s_lc_amend_hdr_lc_amd_no
	
	Const S369_I2_lc_amd_no = 0    'imp s_lc_amend_hdr
    Const S369_I2_amend_dt = 1
    Const S369_I2_amend_req_dt = 2
    Const S369_I2_at_loc_amt = 3
    Const S369_I2_at_xch_rate = 4
    Const S369_I2_at_expiry_dt = 5
    Const S369_I2_at_latest_ship_dt = 6
    Const S369_I2_at_trnshp_flag = 7
    Const S369_I2_at_partial_ship_flag = 8
    Const S369_I2_at_transfer_flag = 9
    Const S369_I2_at_transport = 10
    Const S369_I2_at_loading_port = 11
    Const S369_I2_at_dischge_port = 12
    Const S369_I2_adv_no = 13
    Const S369_I2_pre_adv_ref = 14
    Const S369_I2_remark = 15
    Const S369_I2_remark2 = 16
    Const S369_I2_lc_kind = 17
    Const S369_I2_cur = 18
    Const S369_I2_inc_amt = 19
    Const S369_I2_dec_amt = 20
    Const S369_I2_ext1_qty = 21
    Const S369_I2_ext2_qty = 22
    Const S369_I2_ext3_qty = 23
    Const S369_I2_ext1_amt = 24
    Const S369_I2_ext2_amt = 25
    Const S369_I2_ext3_amt = 26
    Const S369_I2_ext1_cd = 27
    Const S369_I2_ext2_cd = 28
    Const S369_I2_ext3_cd = 29


	
    On Error Resume Next                                                             
    Err.Clear 


	lgIntFlgMode = CInt(Request("txtFlgMode"))									
	
    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"		
    End If
	
	Redim I2_s_lc_amend_hdr(29)				
		
	I2_s_lc_amend_hdr(S369_I2_lc_amd_no) = UCase(Trim(Request("txtLCAmdNo1")))
	I3_s_lc_hdr_lc_no =  Trim(Request("txtLCNo"))
	I2_s_lc_amend_hdr(S369_I2_amend_dt) = UNIConvDate(Trim(Request("txtAmendDt")))	
	
	If Trim(Request("txtRadio")) = "I" Then
		I2_s_lc_amend_hdr(S369_I2_inc_amt) = UNIConvNum(Request("txtAmendAmt"),0)
	ElseIf Trim(Request("txtRadio")) = "D" Then
		I2_s_lc_amend_hdr(S369_I2_dec_amt) = UNIConvNum(Request("txtAmendAmt"),0)
	End If	

	I2_s_lc_amend_hdr(S369_I2_at_xch_rate) = UNIConvNum(Request("txtAtXchRate"),0)
	I2_s_lc_amend_hdr(S369_I2_at_loc_amt)  = UNIConvNum(Request("txtAtLocAmt"),0)
	I2_s_lc_amend_hdr(S369_I2_cur)  = Trim(Request("txtAtCurrency1"))		
	I2_s_lc_amend_hdr(S369_I2_at_expiry_dt) = UNIConvDate(Request("txtAtExpireDt"))	
	I2_s_lc_amend_hdr(S369_I2_at_latest_ship_dt) = UNIConvDate(Request("txtAtLatestShipDt"))	
	
	I2_s_lc_amend_hdr(S369_I2_at_partial_ship_flag) = Trim(Request("rdoAtPartialShip"))
	I2_s_lc_amend_hdr(S369_I2_adv_no) = Trim(Request("txtAdvNo"))
	I2_s_lc_amend_hdr(S369_I2_pre_adv_ref) = Trim(Request("txtPreAdvRef"))
	I2_s_lc_amend_hdr(S369_I2_remark) = Trim(Request("txtRef"))
	I2_s_lc_amend_hdr(S369_I2_lc_kind) = "L"	
	
	
	Set PS4G211 = Server.CreateObject("PS4G211.cSLcAmendHdrSvr")	
	
	if CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	end If
	
    E2_s_lc_amend_hdr_lc_amd_no =  PS4G211.S_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection, iCommandSent  ,I2_s_lc_amend_hdr ,I3_s_lc_hdr_lc_no )
	
	if CheckSYSTEMError(Err,True) = True Then 
		set PS4G211 = Nothing		
		Exit Sub
	end If	

	Response.Write "<Script language=vbs> " & vbCr   	
	If E2_s_lc_amend_hdr_lc_amd_no <> "" Then
		Response.Write " Parent.frm1.txtLCAmdNo.value     = """ & ConvSPChars(E2_s_lc_amend_hdr_lc_amd_no)	 & """" & vbCr 	            	
		
		
    End If
	Response.Write " Parent.DbSaveOk "	 & vbCr 
	Response.Write "</Script> "			 & vbCr      
    
	set PS4G211 = Nothing	
	
	
End Sub

Sub SubBizDelete()

	Dim iCommandSent

	Dim PS4G211
	
	Dim I2_s_lc_amend_hdr
	Dim I3_s_lc_hdr_lc_no
'	Dim E2_s_lc_amend_hdr_lc_amd_no
	
	Const S369_I2_lc_amd_no = 0    'imp s_lc_amend_hdr
    Const S369_I2_amend_dt = 1
    Const S369_I2_amend_req_dt = 2
    Const S369_I2_at_loc_amt = 3
    Const S369_I2_at_xch_rate = 4
    Const S369_I2_at_expiry_dt = 5
    Const S369_I2_at_latest_ship_dt = 6
    Const S369_I2_at_trnshp_flag = 7
    Const S369_I2_at_partial_ship_flag = 8
    Const S369_I2_at_transfer_flag = 9
    Const S369_I2_at_transport = 10
    Const S369_I2_at_loading_port = 11
    Const S369_I2_at_dischge_port = 12
    Const S369_I2_adv_no = 13
    Const S369_I2_pre_adv_ref = 14
    Const S369_I2_remark = 15
    Const S369_I2_remark2 = 16
    Const S369_I2_lc_kind = 17
    Const S369_I2_cur = 18
    Const S369_I2_inc_amt = 19
    Const S369_I2_dec_amt = 20
    Const S369_I2_ext1_qty = 21
    Const S369_I2_ext2_qty = 22
    Const S369_I2_ext3_qty = 23
    Const S369_I2_ext1_amt = 24
    Const S369_I2_ext2_amt = 25
    Const S369_I2_ext3_amt = 26
    Const S369_I2_ext1_cd = 27
    Const S369_I2_ext2_cd = 28
    Const S369_I2_ext3_cd = 29
	
	Redim I2_s_lc_amend_hdr(29)

	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
	If Request("txtLCAmdNo") = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	iCommandSent = "DELETE"	
	I2_s_lc_amend_hdr(S369_I2_lc_amd_no) = Trim(Request("txtLCAmdNo"))
	I3_s_lc_hdr_lc_no =  Trim(Request("txtLCNo"))
		
	Set PS4G211 = Server.CreateObject("PS4G211.cSLcAmendHdrSvr")	
	
	if CheckSYSTEMError(Err,True) = True Then 
		Exit Sub
	end If
	
    Call PS4G211.S_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection, iCommandSent  ,I2_s_lc_amend_hdr ,I3_s_lc_hdr_lc_no )
	
	if CheckSYSTEMError(Err,True) = True Then 
		set PS4G211 = Nothing		
		Exit Sub
	end If	

	Response.Write "<Script language=vbs> " & vbCr  
	Response.Write " Parent.DbDeleteOk "	& vbCr  
	Response.Write "</Script> "				& vbCr      
    
	set PS4G211 = Nothing	
		
End Sub
%>


