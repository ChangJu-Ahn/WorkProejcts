<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7105mb1(�����ڻ꺯������-�Ű�/���)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2001/06/02
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	Err.Clear
	On Error Resume Next														'��: 

	Dim LngMaxRow
	Dim lgCurrency
	Dim lgCurrencyAcq
	Dim lgBlnFlgChgValue, lgOpModeCRUD, lgLngMaxRow

    Call HideStatusWnd                                                               '��: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
'    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    'Single
'    lgPrevNext        = Request("txtPrevNext")                                       '��: "P"(Prev search) "N"(Next search)    
    'Multi SpreadSheet
    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
			call SubBizQuery() 
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

	Response.End    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim I1_a_asset_chg_item_chg_seq 
    Dim I2_a_asset_chg 
    Dim E1_b_minor_minor_nm 
    Dim E2_b_biz_area 
    Dim E3_b_biz_area 
    Dim E4_a_asset_master 
    Dim E5_b_acct_dept 
    Dim E6_a_asset_chg_item_chg_seq 
    Dim E7_a_asset_chg 
    Dim E8_b_biz_partner 
    Dim E9_b_acct_dept 
    Dim E10_b_acct_dept 
    Dim EG1_export_group
    Dim E11_b_tax_biz_area
    '20030301	�̼��ݰ��� 
    Dim E12_a_acct
    '�������߰� 
    dim E13_a_acct
    
    Dim iPAAG035

	Dim IntRows, StrNextKey,  strData
                	
	'  View Name : import a_asset_chg
	Const A509_I2_chg_fg = 0
	Const A509_I2_chg_no = 1

	 '  EXPORTS View ��� 

	'  View Name : exp_fr b_biz_area
	Const A509_E2_biz_area_cd = 0
	Const A509_E2_biz_area_nm = 1

	'  View Name : exp_to b_biz_area
	Const A509_E3_biz_area_cd = 0
	Const A509_E3_biz_area_nm = 1

	'  View Name : exp_master a_asset_master
	Const A509_E4_asst_no = 0
	Const A509_E4_asst_nm = 1
	Const A509_E4_reg_dt = 2
	Const A509_E4_doc_cur = 3
	Const A509_E4_xch_rate = 4
	Const A509_E4_acq_amt = 5
	Const A509_E4_acq_loc_amt = 6
	Const A509_E4_acq_qty = 7
	Const A509_E4_inv_qty = 8

	'  View Name : exp_master b_acct_dept
	Const A509_E5_dept_cd = 0
	Const A509_E5_dept_nm = 1

	'  View Name : export_next a_asset_chg_item
	Const A509_E6_chg_seq = 0

	'  View Name : export a_asset_chg
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

	'  View Name : export b_biz_partner
	Const A509_E8_bp_cd = 0
	Const A509_E8_bp_nm = 1
	'  View Name : export_to b_acct_dept
	Const A509_E9_org_change_id = 0
	Const A509_E9_dept_cd = 1
	Const A509_E9_dept_nm = 2
	'  View Name : export_from b_acct_dept
	Const A509_E10_org_change_id = 0
	Const A509_E10_dept_cd = 1
	Const A509_E10_dept_nm = 2
	
	'  Group Name : export_group
	'  View Name : export_itm_item b_bank_acct
	Const A509_EG1_E1_b_bank_acct_bank_acct_no = 0
	'  View Name : export_itm_item a_asset_chg_item
	Const A509_EG1_E2_a_asset_chg_item_chg_seq = 1
	Const A509_EG1_E2_a_asset_chg_item_paym_type = 2
	Const A509_EG1_E2_a_asset_chg_item_paym_amt = 3
	Const A509_EG1_E2_a_asset_chg_item_paym_loc_amt = 4
	Const A509_EG1_E2_a_asset_chg_item_note_no = 5
	Const A509_EG1_E2_b_minor_nm = 6
	
    '  View Name : export b_tax_biz_area		
    Const A509_E11_tax_biz_area_cd = 0
    Const A509_E11_tax_biz_area_nm = 1
	
	'20030301	�̼��ݰ��� 
	Const A509_E12_ar_ap_acct_cd = 0
	Const A509_E12_ar_ap_acct_nm = 1

    '20030430	�������߰� 
    Const A509_E13_exp_acct_cd = 0
    Const A509_E13_exp_acct_nm = 1
    

     
    ReDim I2_a_asset_chg(A509_I2_chg_no)
    ReDim E2_b_biz_area (A509_E2_biz_area_nm)
    ReDim E3_b_biz_area (A509_E3_biz_area_nm)
    ReDim E4_a_asset_master (A509_E4_inv_qty)
    ReDim E5_b_acct_dept (A509_E5_dept_nm)
    ReDim E7_a_asset_chg (A509_E7_asst_chg_seq)
    ReDim E8_b_biz_partner (A509_E8_bp_nm)
    ReDim E9_b_acct_dept (A509_E9_dept_nm)
    ReDim E10_b_acct_dept (A509_E10_dept_nm)
    ReDim E11_b_tax_biz_area(A509_E11_tax_biz_area_nm)
	'20030301	�̼��ݰ��� 
	ReDim E12_a_acct(A509_E12_ar_ap_acct_nm)
    '20030430	�������߰� 
    Redim E13_a_acct(A509_E13_exp_acct_nm)

	' -- ���Ѱ����߰� 
	Const A312_I3_a_data_auth_data_BizAreaCd = 0
	Const A312_I3_a_data_auth_data_internal_cd = 1
	Const A312_I3_a_data_auth_data_sub_internal_cd = 2
	Const A312_I3_a_data_auth_data_auth_usr_id = 3

	Dim I3_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A312_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I3_a_data_auth(A312_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
		
    Err.Clear                                                                        '��: Clear Error status
    On Error Resume Next                                                             '��: Protect system from crashing

  '********************************************************  
  '                        Query
  '********************************************************  
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
   I2_a_asset_chg(A509_I2_chg_no)=Trim(Request("txtChgNo")) 
    
    '-----------------------
    'Com action area
    '-----------------------     
    call iPAAG035.AS0039_LOOKUP_SVR( gStrGloBalCollection , I1_a_asset_chg_item_chg_seq , I2_a_asset_chg , E1_b_minor_minor_nm  ,  E2_b_biz_area , _
                                     E3_b_biz_area , E4_a_asset_master , E5_b_acct_dept , E6_a_asset_chg_item_chg_seq , E7_a_asset_chg , _
                                     E8_b_biz_partner , E9_b_acct_dept , E10_b_acct_dept, EG1_export_group,  E11_b_tax_biz_area, E12_a_acct, E13_a_acct, I3_a_data_auth) 
  
	'----------------------------------------------
	'Com action result check area(OS,internal)
	'----------------------------------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 

	'-----------------------
	'Result data display area
	'----------------------- 
	lgCurrency = ConvSPChars(E4_a_asset_master(A509_E4_doc_cur))
	lgCurrencyAcq = ConvSPChars(E7_a_asset_chg(A509_E7_doc_cur))
	
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "       	strShowFg=""T""" & vbCr          
	Response.Write "		strChgFg	   = """ & E7_a_asset_chg(A509_E7_chg_fg) &							"""" & vbCr                            '�������� 		 
	Response.Write "    if strChgFg <> ""03"" and strChgFg <> ""04"" then " & vbCr          
	Response.Write "       	IntRetCD = Parent.DisplayMsgBox(""117915"",""X"",""X"",""X"") " & vbCr ''�ڻ������� �ƴմϴ�.         
	Response.Write "       	Parent.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "		strShowFg=""F""	" & vbCr          	
	Response.Write "    end if	" & vbCr          	

	Response.Write "    if strShowFg=""T"" then " & vbCr          
	Response.Write "With parent.frm1						" & vbCr

	
	''''''''''''''''''''''''''''''''
	'  The Part for Asset master
	''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''
	Response.Write "  .txtAsstNo.value     = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_no)) & 				"""" & vbCr
	Response.Write "  .txtAsstNm.value	   = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_nm)) &				"""" & vbCr
	Response.Write "  .fpDateTime1.text    = """ & UNIDateClientFormat(E4_a_asset_master(A509_E4_reg_dt)) &		"""" & vbCr'�ڻ�������� 

	Response.Write "  .txtAcctDeptNm.value = """ & ConvSPChars(E5_b_acct_dept(A509_E5_dept_nm)) &					"""" & vbCr
	Response.Write "  .txtAcqQty.text     = """ & E4_a_asset_master(A509_E4_acq_qty) &							"""" & vbCr
	Response.Write "  .txtInvQty.text     = """ & E4_a_asset_master(A509_E4_inv_qty) &							"""" & vbCr
	
	Response.Write "  .txtCur.value		   = """ & UCase(Trim(lgCurrency)) & 										"""" & vbCr
	Response.Write "  .txtXchRt.text      = """ & E7_a_asset_chg(A509_E7_xch_rate) &							"""" & vbCr
	
	Response.Write "  .txtAcqAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E4_a_asset_master(A509_E4_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr
	Response.Write "  .txtAcqLocAmt.text  = """ & UNIConvNumDBToCompanyByCurrency(E4_a_asset_master(A509_E4_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbCr
	
	''''''''''''''''''''''''''''''''
	'  The Part for Asset Change
	''''''''''''''''''''''''''''''''			
	Response.Write "  .txtChgNo.value      = """ & ConvSPChars(E7_a_asset_chg(A509_E7_chg_no)) &				"""" & vbCr                      '����ȣ         
	Response.Write "  .txtDeptCd.Value     = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_cd)) &			"""" & vbCr                  'ȸ��μ�        
	Response.Write "  .txtDeptCd.Value     = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_cd)) &			"""" & vbCr                  'ȸ��μ�        
	Response.Write "  .hORGchangeID.value = """ & ConvSPChars(E10_b_acct_dept(A509_E10_org_change_id)) &					"""" & vbCr	

	Response.Write "  .txtDeptNm.Value     = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_nm)) &			"""" & vbCr                   'ȸ��μ���	    	    	    
	Response.Write "  .fpDateTime2.text    = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_chg_dt)) &		"""" & vbCr   '��������		          
	Response.Write "		strChgFg	   = """ & E7_a_asset_chg(A509_E7_chg_fg) &							"""" & vbCr                            '�������� 		 
	Response.Write "   .txtChgQty     = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_chg_qty),   ggQty.DecPoint, 0) &					"""" & vbCr                             'ȯ��         
	Response.Write "   .txtBpcd.value        = """ & ConvSPChars(E8_b_biz_partner(A509_E8_bp_cd)) &			"""" & vbCr
	Response.Write "   .txtBpNm.value        = """ & ConvSPChars(E8_b_biz_partner(A509_E8_bp_nm)) &			"""" & vbCr    		 	    
	Response.Write "   .txtDocCur.value      = """ & UCase(Trim(lgCurrencyAcq ))			&				"""" & vbCr                                 '�ŷ���ȭ 
	Response.Write "   .txtXchrate.text     = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_xch_rate),   ggExchRate.DecPoint, 0) &					"""" & vbCr                             'ȯ��         
	Response.Write "	.txtTotalAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_chg_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr  
	Response.Write "	.txtTotalLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_chg_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr   
	
	'20030301	�̼��ݰ��� 
	Response.Write "   .txtArAcctCd.value        = """ & ConvSPChars(E12_a_acct(A509_E12_ar_ap_acct_cd)) &		"""" & vbCr
	Response.Write "   .txtArAcctNm.value        = """ & ConvSPChars(E12_a_acct(A509_E12_ar_ap_acct_nm)) &		"""" & vbCr
	
	Response.Write "	.txtApAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_ar_ap_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &					"""" & vbCr   
	Response.Write "	.txtApLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_ar_ap_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &		"""" & vbCr  
	
	Response.Write "	.txtDeprAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_depr_tot_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &		"""" & vbCr  
	Response.Write "	.txtDeprLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_depr_tot_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &		"""" & vbCr  
	Response.Write "		amt1				= """ & E7_a_asset_chg(A509_E7_ar_ap_amt) &							"""" & vbCr  	
	Response.Write "	if amt1 <> 0 then " & vbCr       '�����ޱ��� �ִ� ��쿡�� �����ޱݸ������� �����ش�. 
	Response.Write "	    .fpDateTime3.text = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_ar_ap_due_dt)) &	"""" & vbCr       'AP ��������       '�������� 
	Response.Write "	end if	" & vbCr
  	Response.Write "     .txtApNo.value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_ar_ap_no)) &				"""" & vbCr
 	Response.Write "	.txtTempGLNo.Value    = """ & ConvSPChars(E7_a_asset_chg(A509_E7_temp_gl_no)) &				"""" & vbCr                                    'TempGL No
	Response.Write "    .txtGLNo.Value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_gl_no)) &					"""" & vbCr                                    'GL No
	Response.Write "    .txtChgDesc.value     = """ & ConvSPChars(E7_a_asset_chg(A509_E7_asset_chg_desc)) &			"""" & vbCr
	Response.Write "	.txtVatType.Value	  = """ & ConvSPChars(Trim(E7_a_asset_chg(A509_E7_vat_type))) &				"""" & vbCr
	Response.Write "	.txtVatTypeNm.Value   = """ & ConvSPChars(E1_b_minor_minor_nm) &							"""" & vbCr
	Response.Write "	.txtVatRate.text	  = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_vat_rate),   ggExchRate.DecPoint, 0) &				"""" & vbCr                               'ȯ�� 
	Response.Write "	.txtVatAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_amt),lgCurrencyAcq,ggAmtOfMoneyNo, "X" , "X") &				"""" & vbCr
	Response.Write "	.txtVatLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_chg(A509_E7_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") &	"""" & vbCr

'10 �� ���� ��ġ �߰� ���� 
	Response.Write "   .txtReportAreaCd.value        = """ & ConvSPChars(E11_b_tax_biz_area(A509_E11_tax_biz_area_cd)) &				"""" & vbCr
	Response.Write "   .txtReportAreaNm.value        = """ & ConvSPChars(E11_b_tax_biz_area(A509_E11_tax_biz_area_nm)) &				"""" & vbCr
	Response.Write "   .fpDateTime4.text = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_issued_dt)) &	"""" & vbCr       'AP ��������       '�������� 

	Response.Write "   End With    		" & vbCr
   
        '************ For Chg Items ********
	Response.Write "	With Parent		" & vbCr  
	Response.Write "	LngMaxRow = .frm1.vspdData.MaxRows		" & vbCr
	Response.Write "	.frm1.vspdData.MaxRows = LngMaxRow		" & vbCr

	if isarray(EG1_export_group) then
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
			strData = strData & Chr(11)  & lgLngMaxRow + IntRows                                
            strData = strData & Chr(11)  & Chr(12) 
        Next
   	Response.Write "	.ggoSpread.Source = .frm1.vspdData             " & vbCr  
	Response.Write "    .ggoSpread.SSShowData  """ & strData &		"""" & vbCr  

    End if   
    	
	Response.Write "    .lgStrPrevKey = """ & StrNextKey &			""" " & vbCr          
'	Response.Write "	.lgNextNo =							"""" " & vbCr  	' ���� Ű �� �Ѱ��� 
'	Response.Write "	.lgPrevNo =							"""" " & vbCr  	' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵǾ� ���� 															'��: ��ȭ�� ���� 

	Response.Write "    if strChgFg = ""03"" then " & vbCr          
	Response.Write "       .frm1.Rb_Sold.checked = true" & vbCr          
	Response.Write "       .DbQueryOk	" & vbCr          
	Response.Write "    elseif strChgFg = ""04"" then " & vbCr          
	Response.Write "       .frm1.Rb_Duse.checked = true  " & vbCr          
	Response.Write "       .DbQueryOk	" & vbCr          
	Response.Write "    else	" & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117915"",""X"",""X"",""X"") " & vbCr         
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "    end if	" & vbCr          	
	Response.Write "	End With		" & vbCr  

	Response.Write "    end if	" & vbCr          	
	
	Response.Write "</Script>		" & vbCr  
	 
    Set iPAAG035 = Nothing															    '��: Unload Comproxy

	Response.End																		'��: Process End
		
end sub
    
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================

Sub SubBizSave()
	Dim iCommandSent
	Dim I1_a_asset_chg 
	Dim I2_b_acct_dept 
	Dim I3_b_biz_partner_bp_cd 
	Dim I4_b_currency_currency 
	Dim I5_b_acct_dept 
	Dim I6_a_asset_chg 
	Dim IG1_import_group 
	Dim I7_a_asset_master_asst_no 
	Dim I8_ief_supplied_select_char 
	Dim iPAAG035
	Dim E2_a_asset_chg 	
    '  View Name : import_org_insrt a_asset_chg
    Const A507_I1_insrt_user_id = 0
    Const A507_I1_insrt_dt = 1
    '  View Name : import_to b_acct_dept
    Const A507_I2_org_change_id = 0
    Const A507_I2_dept_cd = 1
    '  View Name : import b_acct_dept
    Const A507_I5_org_change_id = 0
    Const A507_I5_dept_cd = 1
    Const A507_I5_internal_cd = 2
    '  View Name : import a_asset_chg
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
	'20030301	�̼��ݰ��� 
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
	'20030430	��������߰� 
    Const A507_I6_exp_acct_cd = 36 

    Const A507_I6_asst_chg_seq = 37  'View Name : import b_tax_biz_area



    '  IMPORTS Group View ��� 
    '  Group Name : import_group
    '  View Name : import_item ief_supplied
    Const A507_IG1_I1_ief_supplied_select_char = 0
    '  View Name : import_item b_bank_acct
    Const A507_IG1_I2_b_bank_acct_bank_acct_no = 1
    '  View Name : import_item a_asset_chg_item
    Const A507_IG1_I3_a_asset_chg_item_chg_seq = 2
    Const A507_IG1_I3_a_asset_chg_item_paym_type = 3
    Const A507_IG1_I3_a_asset_chg_item_paym_amt = 4
    Const A507_IG1_I3_a_asset_chg_item_paym_loc_amt = 5
    
    reDim I1_a_asset_chg (A507_I1_insrt_dt)
	reDim I2_b_acct_dept (A507_I2_dept_cd)
	reDim I5_b_acct_dept (A507_I5_internal_cd)
	reDim I6_a_asset_chg (A507_I6_asst_chg_seq)

	' -- ���Ѱ����߰� 
	Const A312_I9_a_data_auth_data_BizAreaCd = 0
	Const A312_I9_a_data_auth_data_internal_cd = 1
	Const A312_I9_a_data_auth_data_sub_internal_cd = 2
	Const A312_I9_a_data_auth_data_auth_usr_id = 3

	Dim I9_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A312_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	'***************************************************************
	'                              SAVE
	'***************************************************************									
    Err.Clear                                                                        '��: Clear Error status
    On Error Resume Next                                                             '��: Protect system from crashing
	
	Dim lgIntFlgMode

	lgIntFlgMode = CInt(Request("txtFlgMode"))									        '��: ����� Create/Update �Ǻ� 
	LngMaxRow    = CInt(Request("txtMaxRows"))	
	
    Set iPAAG035 = Server.CreateObject("PAAG035.cAAS0031MngSvr") 
    
	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 
'    gChangeOrgId = GetGlobalInf("gChangeOrgId")
    gChangeOrgId =REquest("hORGCHANGEID")
    '-----------------------
    'Data manipulate area
    '-----------------------
    I4_b_currency_currency   = gCurrency							   '�ڱ���ȭ        
    I7_a_asset_master_asst_no    = Trim(Request("txtAsstNo"))       
    I5_b_acct_dept(A507_I5_dept_cd)       = Trim(Request("txtDeptCd"))
    I5_b_acct_dept(A507_I5_org_change_id)  = gChangeOrgId    
	I3_b_biz_partner_bp_cd       = Trim(Request("txtBpCd"))	           '�ŷ�ó 
    
    I6_a_asset_chg(A507_I6_chg_no)	    = Trim(Request("txtChgNo"))   
    I6_a_asset_chg(A507_I6_chg_dt)		= UNIConvDate(Request("txtChgDt"))        '��������    		    
    I6_a_asset_chg(A507_I6_chg_fg)        = Request("radio1")				       '��������	
    I6_a_asset_chg(A507_I6_chg_qty)        =UNIConvNum(Request("txtChgQty"),0)		'�Ű�/������	
	I6_a_asset_chg(A507_I6_doc_cur)       = UCase(Request("txtDocCur"))                   '�ŷ���ȭ 
	
	if UCase(Request("txtDocCur")) = gCurrency then        
		I6_a_asset_chg(A507_I6_xch_rate)  = 1
	else
		I6_a_asset_chg(A507_I6_xch_rate)  = UNIConvNum(Request("txtXchRate"),0)        'ȯ�� 
	end if			
	
	I6_a_asset_chg(A507_I6_chg_amt)      = UNIConvNum(Request("txtTotalAmt"),0)
	I6_a_asset_chg(A507_I6_chg_loc_amt)   = UNIConvNum(Request("txtTotalLocAmt"),0) 
	
	'20030301	�̼��ݰ��� 
	I6_a_asset_chg(A507_I6_ar_ap_acct_cd)  = Trim(Request("txtArAcctCd"))
	
	I6_a_asset_chg(A507_I6_ar_ap_due_dt)  = UNIConvDate(Request("txtDueDt"))             'AR/AP �������� 		
	I6_a_asset_chg(A507_I6_ar_ap_no)       = Trim(Request("txtApNo"))            'AMEND ���� �ʿ�         		 
	I6_a_asset_chg(A507_I6_temp_gl_no)		= Trim(Request("txtTempGLNo"))				 
	I6_a_asset_chg(A507_I6_gl_no)			= Trim(Request("txtGlNo"))				
	I6_a_asset_chg(A507_I6_asset_chg_desc) = Request("txtChgDesc")                 '��������       
	I6_a_asset_chg(A507_I6_vat_io_fg)	=	"O"
	I6_a_asset_chg(A507_I6_vat_type)	=	UCase(Trim(Request("txtVatType"))) 
	I6_a_asset_chg(A507_I6_vat_rate)	=	UNIConvNum(Request("txtVatRate"),0)
	I6_a_asset_chg(A507_I6_vat_amt)	=	UNIConvNum(Request("txtVatAmt"),0)		
	I6_a_asset_chg(A507_I6_vat_loc_amt)	=	UNIConvNum(Request("txtVatLocAmt"),0)		
	
	If Request("txtIssuedDt") <>"" Then
		I6_a_asset_chg(A507_I6_issued_dt)	=	UNIConvDate(Request("txtIssuedDt"))   ' 10�� ���� ��ġ �߰� 
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
		Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT) '��: you must release this line if you change msg into code
    End If

	'-----------------------
	'Com Action Area
	'-----------------------


	 E2_a_asset_chg = iPAAG035.AS0031_MANAGE_SVR( gStrGloBalCollection ,iCommandSent , I2_b_acct_dept ,I3_b_biz_partner_bp_cd , _
	                                            I4_b_currency_currency ,I5_b_acct_dept ,I6_a_asset_chg ,Request("txtSpread"),I7_a_asset_master_asst_no , _
	                                            I8_ief_supplied_select_char,I9_a_data_auth ) 

	'-----------------------
	'DB Error
	'-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
		Exit Sub
    End If 

   Set iPAAG035 = Nothing                                                  '��: Unload 

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtChgNo.Value=  """ & ConvSPChars(E2_a_asset_chg) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵ� ���� 															'��: ��ȭ�� ���� 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	Dim E2_a_asset_chg
	Dim I6_a_asset_chg
	Dim iPAAG035
    Const A507_I6_chg_no = 0
    Const A507_I6_chg_dt = 1
    Const A507_I6_chg_fg = 2
    Const A509_E7_vat_loc_amt = 26    

	reDim I6_a_asset_chg (A509_E7_vat_loc_amt)
	
	' -- ���Ѱ����߰� 
	Const A312_I9_a_data_auth_data_BizAreaCd = 0
	Const A312_I9_a_data_auth_data_internal_cd = 1
	Const A312_I9_a_data_auth_data_sub_internal_cd = 2
	Const A312_I9_a_data_auth_data_auth_usr_id = 3

	Dim I9_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A312_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A312_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))	

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	'***************************************************************
	'                              DELETE
	'***************************************************************

    Err.Clear                                                                        '��: Clear Error status
    On Error Resume Next                                                             '��: Protect system from crashing

    If Request("txtChgNo") = "" Then    	'��: ������ ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'���� ���ǰ��� ����ֽ��ϴ�!
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

    I6_a_asset_chg(A507_I6_chg_no)   = Trim(Request("txtChgNo"))   
    I6_a_asset_chg(A507_I6_chg_fg)   = CSTR(Request("Radio1"))    
    
     E2_a_asset_chg = iPAAG035.AS0031_MANAGE_SVR( gStrGloBalCollection ,"DELETE" , , , , , I6_a_asset_chg ,,,,I9_a_data_auth)         
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 

    Set iPAAG035 = Nothing                                                   '��: Unload 
    
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	Call parent.DbDeleteOk()		" & vbCr
	Response.Write "</Script>		" & vbCr 
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
	
%>
