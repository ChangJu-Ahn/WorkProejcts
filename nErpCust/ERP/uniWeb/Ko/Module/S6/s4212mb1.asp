<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4212mb1.asp																*
'*  4. Program Name         : 통관내역등록																*
'*  5. Program Desc         : 통관내역등록																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2000/04/11																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/04/17 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%												

On Error Resume Next
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iLngRow
Dim intGroupCount

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 
lgStrPrevKey = Request("lgStrPrevKey")


Select Case strMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	Err.Clear																
															
	Dim iPS6G219		
    Dim I1_s_cc_hdr_cc_no	
    Dim EG1_s_cc_hdr
    Dim pvCommandSent
        
    Const E1_cc_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_cc_hdr
    Const E1_iv_no = 1
    Const E1_iv_dt = 2
    Const E1_so_no = 3
    Const E1_manufacturer = 4
    Const E1_agent = 5
    Const E1_return_appl = 6
    Const E1_return_office = 7
    Const E1_reporter = 8
    Const E1_ed_no = 9
    Const E1_ed_dt = 10
    Const E1_ed_type = 11
    Const E1_inspect_type = 12
    Const E1_ep_type = 13
    Const E1_export_type = 14
    Const E1_incoterms = 15
    Const E1_pay_meth = 16
    Const E1_pay_dur = 17
    Const E1_dischge_cntry = 18
    Const E1_loading_port = 19
    Const E1_loading_cntry = 20
    Const E1_dischge_port = 21
    Const E1_vessel_nm = 22
    Const E1_transport = 23
    Const E1_trans_form = 24
    Const E1_inspect_req_dt = 25
    Const E1_device_plce = 26
    Const E1_lc_doc_no = 27
    Const E1_lc_amend_seq = 28
    Const E1_lc_no = 29
    Const E1_lc_type = 30
    Const E1_lc_open_dt = 31
    Const E1_open_bank = 32
    Const E1_gross_weight = 33
    Const E1_weight_unit = 34
    Const E1_tot_packing_cnt = 35
    Const E1_packing_type = 36
    Const E1_ship_fin_dt = 37
    Const E1_cur = 38
    Const E1_doc_amt = 39
    Const E1_fob_doc_amt = 40
    Const E1_xch_rate = 41
    Const E1_loc_amt = 42
    Const E1_fob_loc_amt = 43
    Const E1_freight_loc_amt = 44
    Const E1_insure_loc_amt = 45
    Const E1_el_doc_no = 46
    Const E1_el_app_dt = 47
    Const E1_ep_no = 48
    Const E1_ep_dt = 49
    Const E1_insp_cert_no = 50
    Const E1_insp_cert_dt = 51
    Const E1_quar_cert_no = 52
    Const E1_quar_cert_dt = 53
    Const E1_recomnd_no = 54
    Const E1_recomnd_dt = 55
    Const E1_trans_method = 56
    Const E1_trans_rep_cd = 57
    Const E1_trans_from_dt = 58
    Const E1_trans_to_dt = 59
    Const E1_customs = 60
    Const E1_final_dest = 61
    Const E1_remark1 = 62
    Const E1_remark2 = 63
    Const E1_remark3 = 64
    Const E1_origin = 65
    Const E1_origin_cntry = 66
    Const E1_usd_xch_rate = 67
    Const E1_biz_area = 68
    Const E1_ref_flag = 69
    Const E1_sts = 70
    Const E1_net_weight = 71
    Const E1_ext1_qty = 72
    Const E1_ext2_qty = 73
    Const E1_ext3_qty = 74
    Const E1_ext1_amt = 75
    Const E1_ext2_amt = 76
    Const E1_ext3_amt = 77
    Const E1_ext1_cd = 78
    Const E1_ext2_cd = 79
    Const E1_ext3_cd = 80
    Const E1_xch_rate_op = 81
    Const E2_bp_cd = 82   
    Const E2_bp_nm = 83    
    Const E3_bp_cd = 84   
    Const E3_bp_nm = 85
    Const E4_sales_grp = 86
    Const E4_sales_grp_nm = 87    
    Const E5_sales_org = 88
    Const E5_sales_org_nm = 89
    Const E6_bp_nm = 90    
    Const E7_bp_nm = 91    
    Const E8_bp_nm = 92    
    Const E9_bp_nm = 93    
    Const E10_bank_nm = 94 
    Const E11_minor_nm = 95
    Const E12_minor_nm = 96
    Const E13_minor_nm = 97
    Const E14_minor_nm = 98
    Const E15_minor_nm = 99
    Const E16_minor_nm = 100
    Const E17_minor_nm = 101
    Const E18_minor_nm = 102
    Const E19_country_nm = 103
    Const E20_minor_nm = 104  
    Const E21_country_nm = 105
    Const E22_minor_nm = 106  
    Const E23_country_nm = 107
    Const E24_minor_nm = 108  
    Const E25_minor_nm = 109  
    Const E26_minor_nm = 110  
    Const E27_minor_nm = 111  
    Const E28_bp_nm = 112     
    Const E29_minor_nm = 113
    
    pvCommandSent = "QUERY"
                															
	I1_s_cc_hdr_cc_no = Trim(Request("txtCCNo"))
	'---------------------------------- C/C Header Data Query ----------------------------------

	Set iPS6G219 = Server.CreateObject("PS6G219.cSLkExportCcHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
       'Exit Sub
    End If   
	  		
	Call iPS6G219.S_LOOKUP_EXPORT_CC_HDR_SVR(gStrGlobalCollection, pvCommandSent, I1_s_cc_hdr_cc_no, _
	                                         EG1_s_cc_hdr)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS6G219 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
       'Exit Sub
    End If   

    Set iPS6G219 = Nothing   
	Dim lgCurrency
	lgCurrency = ConvSPChars(EG1_s_cc_hdr(E1_cur))
		

	Response.Write "<Script language=vbs>  " & vbCr   			    

	Response.Write " Parent.frm1.txtCurrency.value   = """ & lgCurrency     & """" & vbCr    
	Response.Write " Parent.CurFormatNumericOCX  " & vbCr   		

	Response.Write " Parent.frm1.txtApplicant.value	= """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_cd))			& """" & vbcr
	Response.Write " Parent.frm1.txtApplicantNm.value	= """ & ConvSPChars(EG1_s_cc_hdr(E2_bp_nm))         & """" & vbcr
	Response.Write " Parent.frm1.txtSONo.value			= """ & ConvSPChars(EG1_s_cc_hdr(E1_so_no))         & """" & vbcr
	Response.Write " Parent.frm1.txtSalesGroup.value	= """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp))     & """" & vbcr
	Response.Write " Parent.frm1.txtSalesGroupNm.value	= """ & ConvSPChars(EG1_s_cc_hdr(E4_sales_grp_nm))  & """" & vbcr
	Response.Write " Parent.frm1.txtLCDocNo.value		= """ & ConvSPChars(EG1_s_cc_hdr(E1_lc_doc_no))     & """" & vbcr
		
	Response.Write " If Len(Parent.frm1.txtLCDocNo.value) Then	" & vbCr   		
	Response.Write " Parent.frm1.txtLCAmendSeq.value	 = """ & ConvSPChars(EG1_s_cc_hdr(E1_lc_amend_seq))     & """" & vbcr
	Response.Write " End If 						" & vbCr   		
		
	Response.Write " Parent.frm1.txtPayTerms.value		= """ & ConvSPChars(EG1_s_cc_hdr(E1_pay_meth))												 & """" & vbcr
	Response.Write " Parent.frm1.txtPayTermsNm.value	= """ & ConvSPChars(EG1_s_cc_hdr(E17_minor_nm))                                             & """" & vbcr
	Response.Write " Parent.frm1.txtIncoTerms.value		= """ & ConvSPChars(EG1_s_cc_hdr(E1_incoterms))                                             & """" & vbcr
	Response.Write " Parent.frm1.txtIncoTermsNm.value	= """ & ConvSPChars(EG1_s_cc_hdr(E16_minor_nm))                                             & """" & vbcr
	Response.Write " Parent.frm1.txtCurrency.value		= """ & ConvSPChars(EG1_s_cc_hdr(E1_cur))                                                   & """" & vbcr

	Response.Write " Parent.frm1.txtDocAmt.value		= """ & UNINumClientFormatByCurrency(EG1_s_cc_hdr(E1_doc_amt),lgCurrency,ggAmtOfMoneyNo)    & """" & vbcr
	Response.Write " Parent.frm1.txtWeightUnit.value	= """ & ConvSPChars(EG1_s_cc_hdr(E1_weight_unit))                                           & """" & vbcr
	Response.Write " Parent.frm1.txtNetWeight.value		= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_net_weight), ggQty.DecPoint, 0)    & """" & vbcr
	Response.Write " Parent.frm1.txtRefFlg.value		= """ & ConvSPChars(EG1_s_cc_hdr(E1_ref_flag))                                              & """" & vbcr
	Response.Write " Parent.frm1.txtHCCNo.value			= """ & ConvSPChars(Request("txtCCNo"))                                                     & """" & vbcr		
	Response.Write " Parent.frm1.txtHXchRate.value		= """ & UNINumClientFormat(EG1_s_cc_hdr(E1_xch_rate), ggExchRate.DecPoint, 0) & """" & vbcr
	Response.Write " Parent.frm1.txtHXchRateOp.value	= """ & ConvSPChars(EG1_s_cc_hdr(E1_xch_rate_op))                                           & """" & vbcr
	Response.Write " Parent.frm1.txtMaxSeq.value		= """ & "0"		& """" & vbcr		
	Response.Write " If Parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then Parent.CurFormatNumSprSheet " & vbCr   
	Response.Write " Call parent.CCHdrQueryOk() " & vbCr   
	Response.Write "</Script>      " & vbCr      

    Dim iPS6G228
    Dim I2_s_cc_hdr
    Dim I1_s_cc_dtl
         
    Const C_SHEETMAXROWS_D  = 100
        
    Dim EG1_exp_grp
    Const EG1_E1_plant_cd = 0    
    Const EG1_E1_plant_nm = 1
    Const EG1_E2_lc_seq = 2    
    Const EG1_E3_lc_no = 3    
    Const EG1_E3_lc_doc_no = 4
    Const EG1_E4_so_seq = 5   
    Const EG1_E5_so_no = 6  
    Const EG1_E6_so_schd_no = 7    
    Const EG1_E7_dn_seq = 8   
    Const EG1_E8_dn_no = 9    
    Const EG1_E9_cc_no = 10   
    Const EG1_E10_cc_seq = 11   
    Const EG1_E10_hs_cd = 12
    Const EG1_E10_qty = 13
    Const EG1_E10_unit = 14
    Const EG1_E10_price = 15
    Const EG1_E10_doc_amt = 16
    Const EG1_E10_loc_amt = 17
    Const EG1_E10_net_weight = 18
    Const EG1_E10_weight_unit = 19
    Const EG1_E10_lan_no = 20
    Const EG1_E10_bl_qty = 21
    Const EG1_E10_biz_area = 22
    Const EG1_E10_mvmt_no = 23
    Const EG1_E10_subctrct_po_no = 24
    Const EG1_E10_subctrct_po_seq = 25
    Const EG1_E10_ext1_qty = 26
    Const EG1_E10_ext2_qty = 27
    Const EG1_E10_ext3_qty = 28
    Const EG1_E10_ext1_amt = 29
    Const EG1_E10_ext2_amt = 30
    Const EG1_E10_ext3_amt = 31
    Const EG1_E10_ext1_cd = 32
    Const EG1_E10_ext2_cd = 33
    Const EG1_E10_ext3_cd = 34
    Const EG1_E11_item_cd = 35    '[CONVERSION INFORMATION]  View Name : exp_item b_item
    Const EG1_E11_item_nm = 36
    Const EG1_E10_tracking_no = 37
    Const EG1_E11_spec  = 38
    Const EG1_E11_packing_qty = 39   
        
    Dim LngLastRow      
    Dim LngMaxRow       
        
    Dim strTemp
    Dim strData
    Dim iStrNextKey
        
    I2_s_cc_hdr = Trim(Request("txtCCNo"))     
        
    If Request("lgStrPrevKey") <> "" then
	  I1_s_cc_dtl = Request("lgStrPrevKey")
    Else
	  I1_s_cc_dtl = 0
    End If
   
	Set iPS6G228 = Server.CreateObject("PS6G228.cSLtExportCcDtlSvr")          
 
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
       'Exit Sub
    End If   

	Call iPS6G228.S_LIST_EXPORT_CC_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I2_s_cc_hdr, _
	                                       I1_s_cc_dtl, EG1_exp_grp)
	                                           	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS6G228 = Nothing	
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write "   Parent.frm1.txtCCNo.focus " & vbCr    
       Response.Write "</Script>      " & vbCr
       Response.End
       'Exit Sub
    End If   

    Set iPS6G228 = Nothing   
                
    LngMaxRow = CLng(Request("txtMaxRows"))										

	For iLngRow = 0 To UBound(EG1_exp_grp,1)
	    If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_cc_seq)) 
           Exit For
        End If  
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E11_item_cd))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E11_item_nm))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_unit))
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E10_qty), ggQty.DecPoint, 0)

        strData = strData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_E10_price), lgCurrency, ggUnitCostNo)
        strData = strData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_E10_doc_amt), lgCurrency, ggAmtOfMoneyNo)
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E10_net_weight), ggQty.DecPoint, 0)
        strData = strData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_E11_packing_qty), ggQty.DecPoint, 0)
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_hs_cd))
        strData = strData & Chr(11) 
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_lan_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E1_plant_cd))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E8_dn_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E7_dn_seq))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E5_so_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E4_so_seq))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E6_so_schd_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_lc_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E3_lc_doc_no))
		strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E2_lc_seq))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_mvmt_no ))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_subctrct_po_no))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_subctrct_po_seq))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_cc_seq))
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E10_tracking_no))		
        strData = strData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, EG1_E11_spec))
        strData = strData & Chr(11) & LngMaxRow + iLngRow
        strData = strData & Chr(11) & Chr(12)
    
    Next            
    
    Response.Write "<Script language=vbs> " & vbCr           
    Response.Write " Parent.frm1.vspdData.ReDraw = False " & vbCr
    Response.Write " Parent.ggoSpread.Source	= Parent.frm1.vspdData " &	 	  vbCr
    Response.Write " Parent.ggoSpread.SSShowData      """ & strData	 & """" & vbCr
    
    Response.Write " Parent.lgStrPrevKey              = """ & iStrNextKey						& """" & vbCr  
    Response.Write " Parent.frm1.txtHCCNo.value       = """ & ConvSPChars(Request("txtCCNo"))   & """" & vbCr                
    Response.Write " Parent.DbQueryOk "															& vbCr   
    Response.Write "</Script> "																	& vbCr      
    
Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음 
		
	Dim iPS6G221																' 수출통관 Detail Save용 Object
    Dim iErrorPosition		 
		 
    Set iPS6G221 = Server.CreateObject("PS6G221.cSExportCcDtlSvr")          

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If   
	Dim reqtxtSpread
	Dim arrRowVal
	Dim count
	reqtxtSpread = Request("txtSpread")
	Call iPS6G221.S_MAINT_EXPORT_CC_DTL_SVR(gStrGlobalCollection, Trim(Request("txtCCNo")), _
	                                        Trim(reqtxtSpread), iErrorPosition)
	                                           	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS6G221 = Nothing
       Response.End
	End If  

    Set iPS6G221 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr      
	Response.Write " Parent.DBSaveOk "		& vbCr   
	Response.Write "</Script> "				& vbCr      													

Case Else
	Response.End
End Select
%>
