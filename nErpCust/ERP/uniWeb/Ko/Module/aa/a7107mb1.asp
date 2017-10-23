<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7105mb1(고정자산변동내역-자본적/수익적지출)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/3/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next														'☜: 

	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	DIM strApDueDt
	Dim amt1 
	Dim amt2 
	Dim amt3 
	Dim amt4 
	DIm strChgFg
	Dim lgCurrency
	Dim lgCurrencyAcq

	Dim lgIntFlgMode
	Dim lgOpModeCRUD, lgLngMaxRow
	Dim LngMaxRow
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Dim  lgBlnFlgChgValue
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read MAX (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
			call SubBizQuery() 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Response.End       
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    dim I1_a_asset_chg_item_chg_seq 
    dim I2_a_asset_chg 
    dim E1_b_minor_minor_nm 
    dim E2_b_biz_area 
    dim E3_b_biz_area 
    dim E4_a_asset_master 
    dim E5_b_acct_dept 
    dim E6_a_asset_chg_item_chg_seq 
    dim E7_a_asset_chg 
    dim E8_b_biz_partner 
    dim E9_b_acct_dept 
    dim E10_b_acct_dept 
    dim EG1_export_group
    dim E11_b_tax_biz_area
    '미지급금계정추가 
    dim E12_a_acct
    '비용계정추가 
    dim E13_a_acct
    
    dim iPAAG035
	                	
	'[CONVERSION INFORMATION]  View Name : import a_asset_chg
	Const A509_I2_chg_fg = 0
	Const A509_I2_chg_no = 1

	'[CONVERSION INFORMATION]  EXPORTS View 상수 
	'[CONVERSION INFORMATION]  View Name : exp_fr b_biz_area
	Const A509_E2_biz_area_cd = 0
	Const A509_E2_biz_area_nm = 1

	'[CONVERSION INFORMATION]  View Name : exp_to b_biz_area
	Const A509_E3_biz_area_cd = 0
	Const A509_E3_biz_area_nm = 1

	'[CONVERSION INFORMATION]  View Name : exp_master a_asset_master
	Const A509_E4_asst_no = 0
	Const A509_E4_asst_nm = 1
	Const A509_E4_reg_dt = 2
	Const A509_E4_doc_cur = 3
	Const A509_E4_xch_rate = 4
	Const A509_E4_acq_amt = 5
	Const A509_E4_acq_loc_amt = 6
	Const A509_E4_acq_qty = 7
	Const A509_E4_inv_qty = 8

	'[CONVERSION INFORMATION]  View Name : exp_master b_acct_dept
	Const A509_E5_dept_cd = 0
	Const A509_E5_dept_nm = 1

	'[CONVERSION INFORMATION]  View Name : export_next a_asset_chg_item
	Const A509_E6_chg_seq = 0

	'[CONVERSION INFORMATION]  View Name : export a_asset_chg
	Const A509_E7_chg_no = 0
	Const A509_E7_chg_dt = 1
	Const A509_E7_chg_fg = 2
	Const A509_E7_doc_cur = 3
	Const A509_E7_xch_rate = 4
	Const A509_E7_chg_amt = 5
	Const A509_E7_chg_loc_amt = 6
	Const A509_E7_chg_qty = 7
	Const A509_E7_ref_no = 8
	Const A509_E7_depr_tot_amt = 9
	Const A509_E7_depr_tot_loc_amt = 10
	Const A509_E7_ar_ap_amt = 11
	Const A509_E7_ar_ap_loc_amt = 12
	Const A509_E7_ar_ap_due_dt = 13
	Const A509_E7_gl_no = 14
	Const A509_E7_ar_ap_no = 15
	Const A509_E7_to_gl_no = 16
	Const A509_E7_temp_gl_no = 17
	Const A509_E7_to_temp_gl_no = 18
	Const A509_E7_asset_chg_desc = 19
	Const A509_E7_vat_io_fg = 20
	Const A509_E7_vat_type = 21
	Const A509_E7_vat_rate = 22
	Const A509_E7_net_amt = 23
	Const A509_E7_net_loc_amt = 24
	Const A509_E7_vat_amt = 25
	Const A509_E7_vat_loc_amt = 26

    Const A509_E7_issued_dt = 27
    Const A509_E7_tax_biz_area_cd = 28  
    Const A509_E7_asst_chg_seq = 29  
	

	'[CONVERSION INFORMATION]  View Name : export b_biz_partner
	Const A509_E8_bp_cd = 0
	Const A509_E8_bp_nm = 1

	'[CONVERSION INFORMATION]  View Name : export_to b_acct_dept
	Const A509_E9_org_change_id = 0
	Const A509_E9_dept_cd = 1
	Const A509_E9_dept_nm = 2

	'[CONVERSION INFORMATION]  View Name : export_from b_acct_dept
	Const A509_E10_org_change_id = 0
	Const A509_E10_dept_cd = 1
	Const A509_E10_dept_nm = 2

	'[CONVERSION INFORMATION]  Group Name : export_group
	'[CONVERSION INFORMATION]  View Name : export_itm_item b_bank_acct
	Const A509_EG1_E1_b_bank_acct_bank_acct_no = 0
	'[CONVERSION INFORMATION]  View Name : export_itm_item a_asset_chg_item
	Const A509_EG1_E2_a_asset_chg_item_chg_seq = 1
	Const A509_EG1_E2_a_asset_chg_item_paym_type = 2
	Const A509_EG1_E2_a_asset_chg_item_paym_amt = 3
	Const A509_EG1_E2_a_asset_chg_item_paym_loc_amt = 4
	Const A509_EG1_E2_a_asset_chg_item_note_no = 5
	Const A509_EG1_E2_b_minor_nm = 6
    
    '[CONVERSION INFORMATION]  View Name : export b_tax_biz_area	
    Const A509_E11_tax_biz_area_cd = 0
    Const A509_E11_tax_biz_area_nm = 1
    
    '20030301	미지급급계정추가 
    Const A509_E12_ar_ap_acct_cd = 0
    Const A509_E12_ar_ap_acct_nm = 1
    
    '20030430	비용계정추가 
    Const A509_E13_exp_acct_cd = 0
    Const A509_E13_exp_acct_nm = 1
    
    
    Redim I2_a_asset_chg(A509_I2_chg_no)
    Redim E2_b_biz_area (A509_E2_biz_area_nm)
    Redim E3_b_biz_area (A509_E3_biz_area_nm)
    Redim E4_a_asset_master (A509_E4_inv_qty)
    Redim E5_b_acct_dept (A509_E5_dept_nm)
    Redim E7_a_asset_chg (A509_E7_asst_chg_seq)
    Redim E8_b_biz_partner (A509_E8_bp_nm)
    Redim E9_b_acct_dept (A509_E9_dept_nm)
    Redim E10_b_acct_dept (A509_E10_dept_nm)
    Redim E11_b_tax_biz_area(A509_E11_tax_biz_area_nm)
    '20030301	미지급급계정추가 
    Redim E12_a_acct(A509_E12_ar_ap_acct_nm)
    '20030430	비용계정추가 
    Redim E13_a_acct(A509_E13_exp_acct_nm)


	' -- 권한관리추가 
	Const A312_I3_a_data_auth_data_BizAreaCd = 0
	Const A312_I3_a_data_auth_data_internal_cd = 1
	Const A312_I3_a_data_auth_data_sub_internal_cd = 2
	Const A312_I3_a_data_auth_data_auth_usr_id = 3

	Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A312_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  '********************************************************  
  '                        Query
  '********************************************************  
    I2_a_asset_chg(A509_I2_chg_no)=Trim(Request("txtChgNo")) 

    Set iPAAG035 = Server.CreateObject("PAAG035.cAAS0039LkUpSvr")    
     
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_a_asset_chg(A509_I2_chg_no)      = Trim(Request("txtChgNo"))   
	    
    '-----------------------
    'Com action area
    '-----------------------        
    call iPAAG035.AS0039_LOOKUP_SVR( gStrGloBalCollection , I1_a_asset_chg_item_chg_seq , I2_a_asset_chg , E1_b_minor_minor_nm  ,  E2_b_biz_area , _
                                     E3_b_biz_area , E4_a_asset_master , E5_b_acct_dept , E6_a_asset_chg_item_chg_seq , E7_a_asset_chg , _
                                     E8_b_biz_partner , E9_b_acct_dept , E10_b_acct_dept, EG1_export_group,  E11_b_tax_biz_area, E12_a_acct, E13_a_acct, I3_a_data_auth ) 

	'----------------------------------------------
	'Com action result check area
	'----------------------------------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 

	lgCurrency = ConvSPChars(E4_a_asset_master(A509_E4_doc_cur))
	lgCurrencyAcq = ConvSPChars(E7_a_asset_chg(A509_E7_doc_cur))
	
	Dim strShowFg, IntRetCD, gIsShowLocal, IntRows, StrNextKey, lgStrPrevKey, strData
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "       	strShowFg=""T""" & vbCr          
	Response.Write "		strChgFg	   = """ & E7_a_asset_chg(A509_E7_chg_fg) &							"""" & vbCr                            '변동구분 		 
	Response.Write "    if strChgFg <> ""01"" and strChgFg <> ""02"" then " & vbCr          
	Response.Write "       	IntRetCD = Parent.DisplayMsgBox(""117914"",""X"",""X"",""X"") " & vbCr ''자산지출이 아닙니다.         
	Response.Write "       	Parent.lgBlnFlgChgValue = False " & vbCr          
'	Response.Write "       	Call Parent.fncnew()" & vbCr          
	Response.Write "		strShowFg=""F""	" & vbCr          	
	Response.Write "    end if	" & vbCr          	

	Response.Write "    if strShowFg=""T"" then " & vbCr          
	Response.Write "With parent.frm1						" & vbCr

	''''''''''''''''''''''''''''''''
	'  The Part for Asset master
	''''''''''''''''''''''''''''''''

	Response.Write "  .txtAsstNo.value     = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_no)) & 				"""" & vbCr
	Response.Write "  .txtAsstNm.value	   = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_nm)) &				"""" & vbCr
	Response.Write "  .fpDateTime1.text    = """ & UNIDateClientFormat(E4_a_asset_master(A509_E4_reg_dt)) &		"""" & vbCr'자산취득일자 
	Response.Write "  .txtAcctDeptNm.value = """ & ConvSPChars(E5_b_acct_dept(A509_E5_dept_nm)) &					"""" & vbCr	
	Response.Write "  .hORGchangeID.value = """ & ConvSPChars(E10_b_acct_dept(A509_E10_org_change_id)) &					"""" & vbCr	

	Response.Write "  .txtAcqQty.text     = """ & E4_a_asset_master(A509_E4_acq_qty) &							"""" & vbCr
	Response.Write "  .txtInvQty.text     = """ & E4_a_asset_master(A509_E4_inv_qty) &							"""" & vbCr	
	Response.Write "  .txtCur.value		   = """ & UCase(Trim(lgCurrency)) & 										"""" & vbCr
	if gIsShowLocal <> "N" then
	Response.Write "  .txtXchRt.text      = """ & E7_a_asset_chg(A509_E7_xch_rate) &							"""" & vbCr
	else
	Response.Write "  .txtXchRt.value      = """ & E7_a_asset_chg(A509_E7_xch_rate) &							"""" & vbCr	
	end if
	
	Response.Write "  .txtAcqAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E4_a_asset_master(A509_E4_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr

	if gIsShowLocal <> "N" then
	Response.Write "  .txtAcqLocAmt.text  = """ & UNIConvNumDBToCompanyByCurrency(E4_a_asset_master(A509_E4_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbCr
	else
	Response.Write "  .txtAcqLocAmt.value  = """ & UNIConvNumDBToCompanyByCurrency(E4_a_asset_master(A509_E4_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbCr
	end if

	''''''''''''''''''''''''''''''''
	'  The Part for Asset Change
	''''''''''''''''''''''''''''''''
	Response.Write "	.txtChgNo.value      = """ & ConvSPChars(E7_a_asset_chg(A509_E7_chg_no)) &				"""" & vbCr					'취득번호 
	Response.Write "	.txtDeptCd.Value     = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_cd)) &			"""" & vbCr                 '회계부서        
	Response.Write "	.txtDeptNm.Value     = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_nm)) &			"""" & vbCr                 '회계부서명	    	    	    
	Response.Write "	.fpDateTime2.text    = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_chg_dt)) &		"""" & vbCr   '변동일자		          
	Response.Write "	.txtBpcd.value        = """ & ConvSPChars(E8_b_biz_partner(A509_E8_bp_cd)) &				"""" & vbCr
	Response.Write "	.txtBpNm.value        = """ & ConvSPChars(E8_b_biz_partner(A509_E8_bp_nm)) &				"""" & vbCr    		 	    
	Response.Write "	.txtDocCur.value      = """ & UCase(Trim(lgCurrencyAcq ))			&					"""" & vbCr                                 '거래통화 
	Response.Write "	.txtTotalAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_chg_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr  
    
	
	Response.Write "	.txtApAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_ar_ap_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &					"""" & vbCr   
If gIsShowLocal <> "N" Then	
	Response.Write "	.txtXchrate.text     = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_xch_rate),   ggExchRate.DecPoint, 0) &					"""" & vbCr                             '환율         
	Response.Write "	.txtTotalLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_chg_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr   
	Response.Write "	.txtApLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_ar_ap_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &		"""" & vbCr  
else
	Response.Write "	.txtXchrate.value     = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_xch_rate),   ggExchRate.DecPoint, 0) &					"""" & vbCr                             '환율         
	Response.Write "	.txtTotalLocAmt.value     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_chg_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr   
	Response.Write "	.txtApLocAmt.value     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_ar_ap_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &		"""" & vbCr  
end if

	Response.Write "	amt1				= """ & E7_a_asset_chg(A509_E7_ar_ap_amt) &							"""" & vbCr  	
	Response.Write "	if amt1 <> 0 then     " & vbCr   '미지급금이 있는 경우에만 미지급금만기일자 보여준다.
	Response.Write "	    .fpDateTime3.text = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_ar_ap_due_dt)) &	"""" & vbCr       'AP 만기일자       '변동일자        
	Response.Write "	end if	" & vbCr
  	Response.Write "    .txtApNo.value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_ar_ap_no)) &				"""" & vbCr          		
 	Response.Write "	.txtTempGLNo.Value    = """ & ConvSPChars(E7_a_asset_chg(A509_E7_temp_gl_no)) &				"""" & vbCr                                    'TempGL No        
	Response.Write "    .txtGLNo.Value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_gl_no)) &					"""" & vbCr                                    'GL No                
	Response.Write "    .txtChgDesc.value     = """ & ConvSPChars(E7_a_asset_chg(A509_E7_asset_chg_desc)) &			"""" & vbCr  
	Response.Write "	.txtVatType.Value	  = """ & ConvSPChars(Trim(E7_a_asset_chg(A509_E7_vat_type))) &				"""" & vbCr  	
	Response.Write "	.txtVatTypeNm.Value   = """ & ConvSPChars(E1_b_minor_minor_nm) &							"""" & vbCr  
If gIsShowLocal <> "N" Then	
	Response.Write "	.txtVatRate.text	  = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_vat_rate),   ggExchRate.DecPoint, 0) &				"""" & vbCr                               '환율			
	Response.Write "	.txtVatAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr  
	Response.Write "	.txtVatLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr   
else
	Response.Write "	.txtVatRate.value	  = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_vat_rate),   ggExchRate.DecPoint, 0) &				"""" & vbCr                               '환율			
	Response.Write "	.txtVatAmt.value        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr  
	Response.Write "	.txtVatLocAmt.value     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr   
end if
'10 월 정기 패치 추가 사항 
	Response.Write "	.txtReportAreaCd.value        = """ & ConvSPChars(E11_b_tax_biz_area(A509_E11_tax_biz_area_cd)) &				"""" & vbCr
	Response.Write "	.txtReportAreaNm.value        = """ & ConvSPChars(E11_b_tax_biz_area(A509_E11_tax_biz_area_nm)) &				"""" & vbCr    		 	    

    '20030301	미지급급계정추가 
'    Response.Write "	.frm1.txtApAcctCd.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_cd))				 & """" & vbCr '미지급금계정            
'    Response.Write "	.frm1.txtApAcctNm.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_nm))				 & """" & vbCr '미지급금계정            
    Response.Write "	.txtApAcctCd.value        = """ & ConvSPChars(E12_a_acct(A509_E12_ar_ap_acct_cd)) &				"""" & vbCr
	Response.Write "	.txtApAcctNm.value        = """ & ConvSPChars(E12_a_acct(A509_E12_ar_ap_acct_nm)) &				"""" & vbCr    		 	    

    '20030430	비용계정추가 
    Response.Write "	.txtExpAcctCd.value        = """ & ConvSPChars(E13_a_acct(A509_E13_exp_acct_cd)) &				"""" & vbCr
	Response.Write "	.txtExpAcctNm.value        = """ & ConvSPChars(E13_a_acct(A509_E13_exp_acct_nm)) &				"""" & vbCr    		 	    

	Response.Write "	.fpDateTime4.text = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_issued_dt)) &	"""" & vbCr       'AP 만기일자       '변동일자        
	Response.Write "   End With    		" & vbCr   


    '************ For Chg Items ********
	Response.Write "	With Parent		" & vbCr  
	Response.Write "	LngMaxRow = .frm1.vspdData.MaxRows		" & vbCr  
	Response.Write "	.frm1.vspdData.MaxRows = LngMaxRow		" & vbCr  

	if IsArray(EG1_export_group) then  
		For IntRows = 0 To ubound(EG1_export_group,1)
            strData = strData & Chr(11)  & EG1_export_group(IntRows,A509_EG1_E2_a_asset_chg_item_chg_seq)
			strData = strData & Chr(11)  & EG1_export_group(IntRows,A509_EG1_E2_a_asset_chg_item_paym_type)
			strData = strData & Chr(11)  & " "    'popup
			strData = strData & Chr(11)  & EG1_export_group(IntRows,A509_EG1_E2_b_minor_nm)    'minor nm
			strData = strData & Chr(11)  & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A509_EG1_E2_a_asset_chg_item_paym_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11)  & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A509_EG1_E2_a_asset_chg_item_paym_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")			
			strData = strData & Chr(11)  & ConvSPChars(EG1_export_group(IntRows,A509_EG1_E1_b_bank_acct_bank_acct_no))    'bank acct
			strData = strData & Chr(11)  & " "    'bak acct pop
			strData = strData & Chr(11)  & ConvSPChars(EG1_export_group(IntRows,A509_EG1_E2_a_asset_chg_item_note_no))
			strData = strData & Chr(11)  & " "	 'note pop																	                'C_ItemPopup
			strData = strData & Chr(11)  & lgLngMaxRow + IntRows+1                                
            strData = strData & Chr(11)  & Chr(12) 
        Next

   		Response.Write "	.ggoSpread.Source = .frm1.vspdData             " & vbCr  
		Response.Write "    .ggoSpread.SSShowData  """ & strData &		"""" & vbCr  

    End if   

	Response.Write "    .lgStrPrevKey = """ & StrNextKey &			""" " & vbCr          
'	Response.Write "	.lgNextNo =							"""" " & vbCr  	' 다음 키 값 넘겨줌 
'	Response.Write "	.lgPrevNo =							"""" " & vbCr  	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 															'☜: 조화가 성공 


	Response.Write "    if strChgFg = ""01"" then " & vbCr          
	Response.Write "       .frm1.Rb_Cpt.checked = true" & vbCr          
	Response.Write "       .DbQueryOk	" & vbCr          
	Response.Write "    elseif strChgFg = ""02"" then " & vbCr          
	Response.Write "       .frm1.Rb_Rve.checked = true  " & vbCr          
	Response.Write "       .DbQueryOk	" & vbCr          
	Response.Write "    else	" & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117914"",""X"",""X"",""X"") " & vbCr ''자산세부내역을 입력하십시오.         
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "    end if	" & vbCr          	
	Response.Write "	End With		" & vbCr  
	
	Response.Write "    end if	" & vbCr          	' show or not
	Response.Write "</Script>		" & vbCr  
	 
    Set iPAAG035 = Nothing															    '☜: Unload 

	Response.End																		'☜: Process End


end sub
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	dim iCommandSent
	dim I1_a_asset_chg 
	dim I2_b_acct_dept 
	dim I3_b_biz_partner_bp_cd 
	dim I4_b_currency_currency 
	dim I5_b_acct_dept 
	dim I6_a_asset_chg 
	dim IG1_import_group 
	dim I7_a_asset_master_asst_no 
	dim I8_ief_supplied_select_char 
	
	Dim iPAAG035
	Dim E2_a_asset_chg 	
	Dim gChangeOrgId
    '[CONVERSION INFORMATION]  View Name : import_org_insrt a_asset_chg
    Const A507_I1_insrt_user_id = 0
    Const A507_I1_insrt_dt = 1
    '[CONVERSION INFORMATION]  View Name : import_to b_acct_dept
    Const A507_I2_org_change_id = 0
    Const A507_I2_dept_cd = 1
    '[CONVERSION INFORMATION]  View Name : import b_acct_dept
    Const A507_I5_org_change_id = 0
    Const A507_I5_dept_cd = 1
    Const A507_I5_internal_cd = 2
    '[CONVERSION INFORMATION]  View Name : import a_asset_chg
    Const A507_I6_chg_no = 0
    Const A507_I6_chg_dt = 1
    Const A507_I6_chg_fg = 2
    Const A507_I6_doc_cur = 3
    Const A507_I6_xch_rate = 4
    Const A507_I6_chg_amt = 5
    Const A507_I6_chg_loc_amt = 6
    Const A507_I6_chg_qty = 7
    Const A507_I6_ref_no = 8
    Const A507_I6_depr_tot_amt = 9
    Const A507_I6_depr_tot_loc_amt = 10
    
    '20030301	미지급금계정추가 
    Const A507_I6_ar_ap_acct_cd = 11
    
    Const A507_I6_ar_ap_amt = 12
    Const A507_I6_ar_ap_loc_amt = 13
    Const A507_I6_ar_ap_due_dt = 14
    Const A507_I6_asset_chg_desc = 15
    Const A507_I6_gl_no = 16
    Const A507_I6_to_gl_no = 17
    Const A507_I6_temp_gl_no = 18
    Const A507_I6_to_temp_gl_no = 19
    Const A507_I6_ar_ap_no = 20
    Const A507_I6_internal_cd = 21
    Const A507_I6_to_internal_cd = 22
    Const A507_I6_insrt_user_id = 23
    Const A507_I6_insrt_dt = 24
    Const A507_I6_updt_user_id = 25
    Const A507_I6_updt_dt = 26
    Const A507_I6_vat_io_fg = 27
    Const A507_I6_vat_type = 28
    Const A507_I6_vat_rate = 29
    Const A507_I6_net_amt = 30
    Const A507_I6_net_loc_amt = 31
    Const A507_I6_vat_amt = 32
    Const A507_I6_vat_loc_amt = 33
    
    Const A507_I6_issued_dt = 34
    Const A507_I6_tax_biz_area_cd = 35  'View Name : import b_tax_biz_area

	'20030430	지출계정추가 
    Const A507_I6_exp_acct_cd = 36 

    Const A507_I6_asst_chg_seq = 37  'View Name : import b_tax_biz_area


    '[CONVERSION INFORMATION]  IMPORTS Group View 상수 
    '[CONVERSION INFORMATION]  Group Name : import_group
    '[CONVERSION INFORMATION]  View Name : import_item ief_supplied
    Const A507_IG1_I1_ief_supplied_select_char = 0
    '[CONVERSION INFORMATION]  View Name : import_item b_bank_acct
    Const A507_IG1_I2_b_bank_acct_bank_acct_no = 1
    '[CONVERSION INFORMATION]  View Name : import_item a_asset_chg_item
    Const A507_IG1_I3_a_asset_chg_item_chg_seq = 2
    Const A507_IG1_I3_a_asset_chg_item_paym_type = 3
    Const A507_IG1_I3_a_asset_chg_item_paym_amt = 4
    Const A507_IG1_I3_a_asset_chg_item_paym_loc_amt = 5
    
    redim I1_a_asset_chg (A507_I1_insrt_dt)
	redim I2_b_acct_dept (A507_I2_dept_cd)
	redim I5_b_acct_dept (A507_I5_internal_cd)
	redim I6_a_asset_chg (A507_I6_asst_chg_seq)
'	redim IG1_import_group (LngMaxRow,A507_IG1_I3_a_asset_chg_item_paym_loc_amt)

	' -- 권한관리추가 
	Const A312_I9_a_data_auth_data_BizAreaCd = 0
	Const A312_I9_a_data_auth_data_internal_cd = 1
	Const A312_I9_a_data_auth_data_sub_internal_cd = 2
	Const A312_I9_a_data_auth_data_auth_usr_id = 3

	Dim I9_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A312_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    

	'***************************************************************
	'                              SAVE
	'***************************************************************									
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operation Mode (CREATE, UPDATE)
	LngMaxRow    = CInt(Request("txtMaxRows"))	

    Set iPAAG035 = Server.CreateObject("PAAG035.cAAS0031MngSvr") 

	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 
	gChangeOrgId = request("hORGCHANGEID")

    '-----------------------
    'Data manipulate area
    '-----------------------
    I4_b_currency_currency   = gCurrency							   '자국통화        
    I7_a_asset_master_asst_no    = Trim(Request("txtAsstNo"))       
    I5_b_acct_dept(A507_I5_dept_cd)       = Trim(Request("txtDeptCd"))
    I5_b_acct_dept(A507_I5_org_change_id)  = gChangeOrgId    
	I3_b_biz_partner_bp_cd       = Trim(Request("txtBpCd"))	           '거래처 	
    I6_a_asset_chg(A507_I6_chg_no)	    = Trim(Request("txtChgNo"))       
    I6_a_asset_chg(A507_I6_chg_dt)		= UNIConvDate(Request("txtChgDt"))        '변동일자    		        
    I6_a_asset_chg(A507_I6_chg_fg)      = Request("radio1")				       '변동구분	
    

    '20030430 비용계정추가 
    If I6_a_asset_chg(A507_I6_chg_fg) = "01" then
		I6_a_asset_chg(A507_I6_exp_acct_cd) = ""
    else
		I6_a_asset_chg(A507_I6_exp_acct_cd) = Request("txtExpAcctCd")
    end if 
    
	I6_a_asset_chg(A507_I6_doc_cur)     = UCase(Request("txtDocCur"))                   '거래통화 
	
	if UCase(Request("txtDocCur")) = gCurrency then        
		I6_a_asset_chg(A507_I6_xch_rate)  = 1
	else
		I6_a_asset_chg(A507_I6_xch_rate)  = UNIConvNum(Request("txtXchRate"),0)        '환율 
	end if			
	
	I6_a_asset_chg(A507_I6_chg_amt)      = UNIConvNum(Request("txtTotalAmt"),0)
	I6_a_asset_chg(A507_I6_chg_loc_amt)  = UNIConvNum(Request("txtTotalLocAmt"),0) 
	
	'20030301	미지급금계정추가 
	I6_a_asset_chg(A507_I6_ar_ap_acct_cd) = Trim(Request("txtApAcctCd")) 
	
	I6_a_asset_chg(A507_I6_ar_ap_due_dt) = UNIConvDate(Request("txtDueDt"))             'AR/AP 만기일자 		
	I6_a_asset_chg(A507_I6_ar_ap_no)     = Trim(Request("txtApNo"))            'AMEND 위해 필요         		 
	I6_a_asset_chg(A507_I6_temp_gl_no)	 = Trim(Request("txtTempGLNo"))				 
	I6_a_asset_chg(A507_I6_gl_no)		 = Trim(Request("txtGlNo"))				
	I6_a_asset_chg(A507_I6_asset_chg_desc) = Request("txtChgDesc")               '변동사유       
	I6_a_asset_chg(A507_I6_vat_io_fg)	=	"I"
	I6_a_asset_chg(A507_I6_vat_type)	=	UCase(Trim(Request("txtVatType"))) 
	I6_a_asset_chg(A507_I6_vat_rate)	=	UNIConvNum(Request("txtVatRate"),0)
	I6_a_asset_chg(A507_I6_vat_amt)		=	UNIConvNum(Request("txtVatAmt"),0)		
	I6_a_asset_chg(A507_I6_vat_loc_amt)	=	UNIConvNum(Request("txtVatLocAmt"),0)		
	
	If Request("txtIssuedDt") <> "" then
		I6_a_asset_chg(A507_I6_issued_dt)	=	UNIConvDate(Request("txtIssuedDt"))   ' 10월 정기 패치 추가 
	End If
	
	I6_a_asset_chg(A507_I6_tax_biz_area_cd)	=	Trim(Request("txtReportAreaCd")) 

    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
		I8_ief_supplied_select_char = "C"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
		I8_ief_supplied_select_char = "U"
    End If
    If LngMaxRow > 30 Then
		Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT) '⊙: you must release this line if you change msg into code
    End If
	'-----------------------
	'Com Action Area
	'-----------------------
	 E2_a_asset_chg = iPAAG035.AS0031_MANAGE_SVR( gStrGloBalCollection ,iCommandSent , I2_b_acct_dept ,I3_b_biz_partner_bp_cd , _
	                                       I4_b_currency_currency ,I5_b_acct_dept ,I6_a_asset_chg ,Request("txtSpread"),I7_a_asset_master_asst_no , _
	                                       I8_ief_supplied_select_char,I9_a_data_auth ) 

 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
	Response.End        
    End If 
    
   Set iPAAG035 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtChgNo.Value=  """ & ConvSPChars(E2_a_asset_chg) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되어 있음 															'☜: 조화가 성공 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  

	Response.End    													   '☜: Process End  	  

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	Dim E2_a_asset_chg
	Dim I6_a_asset_chg
	Dim iPAAG035
    Const A509_I6_chg_no = 0
    Const A509_I6_chg_dt = 1
    Const A509_I6_chg_fg = 2
    Const A509_E7_vat_loc_amt = 26    

	ReDim I6_a_asset_chg (A509_E7_vat_loc_amt)

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	' -- 권한관리추가 
	Const A312_I3_a_data_auth_data_BizAreaCd = 0
	Const A312_I3_a_data_auth_data_internal_cd = 1
	Const A312_I3_a_data_auth_data_sub_internal_cd = 2
	Const A312_I3_a_data_auth_data_auth_usr_id = 3

	Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A312_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	    
	'***************************************************************
	'                              DELETE
	'***************************************************************
    Err.Clear                                                                        '☜: Clear Error status
    On Error Resume Next                                                             '☜: Protect system from crashing

    If Request("txtChgNo") = "" Then    	'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'삭제 조건값이 비어있습니다!
		Response.End 
	End If
    
    Set iPAAG035 = Server.CreateObject("PAAG035.cAAS0031MngSvr") 

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 
    
    I6_a_asset_chg(A509_I6_chg_no)   = Trim(Request("txtChgNo"))   
    I6_a_asset_chg(A509_I6_chg_fg)   = Request("Radio1")
            
    E2_a_asset_chg = iPAAG035.AS0031_MANAGE_SVR( gStrGloBalCollection ,"DELETE" , , , , , I6_a_asset_chg,,,,I3_a_data_auth ) 
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 
    
    Set iPAAG035 = Nothing                                                   '☜: Unload Comproxy
    
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	Call parent.DbDeleteOk()		" & vbCr
	Response.Write "</Script>		" & vbCr 
	
	Response.End    													   '☜: Process End  	  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub


%>
