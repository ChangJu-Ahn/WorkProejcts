<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7109mb1(고정자산부서이동등록)
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
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Err.Clear
	On Error Resume Next														'☜: 

'	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	DIM strApDueDt
	Dim amt1 
	Dim amt2 
	Dim amt3 
	Dim amt4 
	DIm strChgFg
	Dim LngMaxRow
'	Dim lgCurrency
'	Dim lgCurrencyAcq

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Dim  lgBlnFlgChgValue, lgOpModeCRUD
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
'	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
'    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
'    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                            무슨 값??      '☜: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
			call SubBizQuery() 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
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
    '미지급금계정추가 
    dim E12_a_acct
    '비용계정추가 
    dim E13_a_acct
    
	Dim iPAAG035
                	
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
'	Const A509_E7_ar_ap_acct_cd = 11
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

    Const A509_E11_tax_biz_area_cd = 0
    Const A509_E11_tax_biz_area_nm = 1
    
    '20030301	미지급급계정추가 
    Const A509_E12_ar_ap_acct_cd = 0
    Const A509_E12_ar_ap_acct_nm = 1
    
    '20030430	비용계정추가 
    Const A509_E13_exp_acct_cd = 0
    Const A509_E13_exp_acct_nm = 1
    
    redim I2_a_asset_chg(A509_I2_chg_no)
    redim E2_b_biz_area (A509_E2_biz_area_nm)
    redim E3_b_biz_area (A509_E3_biz_area_nm)
    redim E4_a_asset_master (A509_E4_inv_qty)
    redim E5_b_acct_dept (A509_E5_dept_nm)
    redim E7_a_asset_chg (A509_E7_asst_chg_seq)
    redim E8_b_biz_partner (A509_E8_bp_nm)
    redim E9_b_acct_dept (A509_E9_dept_nm)	'to dept
    redim E10_b_acct_dept (A509_E10_dept_nm) ' from dept
    redim E11_b_tax_biz_area (A509_E11_tax_biz_area_nm) ' from dept
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
                                      E8_b_biz_partner , E9_b_acct_dept , E10_b_acct_dept, EG1_export_group, E11_b_tax_biz_area, E12_a_acct, E13_a_acct , I3_a_data_auth)
  
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
	Dim strShowFg, IntRetCD
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "       	strShowFg=""T""" & vbCr          
	Response.Write "		strChgFg	   = """ & E7_a_asset_chg(A509_E7_chg_fg) &							"""" & vbCr                            '변동구분 		 
	Response.Write "    if strChgFg <> ""05""  then " & vbCr          
	Response.Write "       	IntRetCD = Parent.DisplayMsgBox(""117916"",""X"",""X"",""X"") " & vbCr ''자산지출이 아닙니다.         
	Response.Write "       	Parent.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "		strShowFg=""F""	" & vbCr          	
	Response.Write "    end if	" & vbCr          	

	Response.Write "    if strShowFg=""T"" then " & vbCr          
	
'	Response.Write "Dim intRetCd						" & vbCr
	Response.Write "With parent.frm1						" & vbCr
	Response.Write "  .txtChgNo.value      = """ & ConvSPChars(E7_a_asset_chg(A509_E7_chg_no)) &				"""" & vbCr                      '취득번호         
	Response.Write "  .txtAsstNo.value     = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_no)) & 				"""" & vbCr
	Response.Write "  .txtAsstNm.value	   = """ & ConvSPChars(E4_a_asset_master(A509_E4_asst_nm)) &				"""" & vbCr
					
	Response.Write "  .txthFrOrgChangeId.Value    = """ & ConvSPChars(E10_b_acct_dept(A509_E10_org_change_id)) &		"""" & vbCr' from dept  
	Response.Write "  .txtFrDeptCd.Value    = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_cd)) &		"""" & vbCr' from dept  
	Response.Write "  .txtFrDeptCdNm.Value    = """ & ConvSPChars(E10_b_acct_dept(A509_E10_dept_nm)) &		"""" & vbCr'  

	Response.Write "  .txthToOrgChangeId.Value    = """ & ConvSPChars(E9_b_acct_dept(A509_E9_org_change_id)) &		"""" & vbCr' from dept  	
	Response.Write "  .txtToDeptCd.Value    = """ & ConvSPChars(E9_b_acct_dept(A509_E9_dept_cd)) &		"""" & vbCr' from dept  
	Response.Write "  .txtToDeptCdNm.Value    = """ & ConvSPChars(E9_b_acct_dept(A509_E9_dept_nm)) &		"""" & vbCr'  
	Response.Write "  .fpDateTime1    = """ & UNIDateClientFormat(E7_a_asset_chg(A509_E7_chg_dt)) &		"""" & vbCr'변동일자 
	Response.Write "		strChgFg	   = """ & E7_a_asset_chg(A509_E7_chg_fg) &							"""" & vbCr                            '변동구분 		 
	Response.Write "   .txtChgQty     = """ & UNINumClientFormat(E7_a_asset_chg(A509_E7_chg_qty),   ggQty.DecPoint, 0) &		"""" & vbCr                    '환율         
	Response.Write "    .txtFrGLNo.Value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_gl_no)) &					"""" & vbCr                            'GL No                
	Response.Write "    .txtToGLNo.Value        = """ & ConvSPChars(E7_a_asset_chg(A509_E7_to_gl_no)) &					"""" & vbCr                             'GL No                
 	Response.Write "	.txtFrTempGLNo.Value    = """ & ConvSPChars(E7_a_asset_chg(A509_E7_temp_gl_no)) &				"""" & vbCr                             'TempGL No         	
 	Response.Write "	.txtToTempGLNo.Value    = """ & ConvSPChars(E7_a_asset_chg(A509_E7_to_temp_gl_no)) &				"""" & vbCr                          'TempGL No         	
 	Response.Write "	.txtFrBizAreaCd.Value    = """ & ConvSPChars(E2_b_biz_area(A509_E2_biz_area_cd)) &				"""" & vbCr                               
 	Response.Write "	.txtFrBizAreaNm.Value    = """ & ConvSPChars(E2_b_biz_area(A509_E2_biz_area_nm)) &				"""" & vbCr             
 	Response.Write "	.txtToBizAreaCd.Value    = """ & ConvSPChars(E3_b_biz_area(A509_E2_biz_area_cd)) &				"""" & vbCr               
 	Response.Write "	.txtToBizAreaNm.Value    = """ & ConvSPChars(E3_b_biz_area(A509_E3_biz_area_nm)) &				"""" & vbCr             
	Response.Write "    .txthChgRefNo.value     = """ & ConvSPChars(E7_a_asset_chg(A509_E7_ref_no)) &			"""" & vbCr  
	Response.Write "    .txtDesc.value     = """ & ConvSPChars(E7_a_asset_chg(A509_E7_asset_chg_desc)) &			"""" & vbCr     
	Response.Write "    if strChgFg <> ""05"" then " & vbCr         	 
	Response.Write "       	IntRetCD = parent.DisplayMsgBox(""117916"",""X"",""X"",""X"") " & vbCr ''자산부서이동 번호가 아닙니다.         
	Response.Write "       	parent.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call parent.fncnew()" & vbCr          
	Response.Write "    else	" & vbCr          
	Response.Write "       parent.DbQueryOk	" & vbCr          
	Response.Write "    end if	" & vbCr          	
	Response.Write "   End With    		" & vbCr    

	Response.Write "    end if	" & vbCr          	
	Response.Write "</Script>		" & vbCr  

    Set iPAAG035 = Nothing															    '☜: Unload Comproxy
	Response.End																		'☜: Process End
		
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
	Dim lgIntFlgMode
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
    Const A507_I6_ar_ap_acct = 11
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

	Err.Clear
	On Error Resume Next														'☜: 
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))									        '☜: 저장시 Create/Update 판별	
    Set iPAAG035 = Server.CreateObject("PAAG035.cAAS0031MngSvr") 
    
	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 	

gChangeOrgId = GetGlobalInf("gChangeOrgId")
    '-----------------------
    'Data manipulate area
    '-----------------------
    I4_b_currency_currency   = gCurrency							   '자국통화        
    I7_a_asset_master_asst_no    = Trim(Request("txtAsstNo"))       
    I6_a_asset_chg(A507_I6_chg_no)	    = Trim(Request("txtChgNo"))   
    I6_a_asset_chg(A507_I6_chg_dt)		= UNIConvDate(Request("txtChgDt"))        '변동일자    		    
    I6_a_asset_chg(A507_I6_chg_qty)        =UNIConvNum(Request("txtChgQty"),0)		'이동수량 

    I5_b_acct_dept(A507_I5_dept_cd)       = Trim(Request("txtFrDeptCd")) ' from 
    I5_b_acct_dept(A507_I5_org_change_id)  = Trim(Request("txthFrOrgChangeId")) ' from  'gChangeOrgId    

    I2_b_acct_dept(A507_I5_dept_cd)       = Trim(Request("txtToDeptCd"))' to
    I2_b_acct_dept(A507_I5_org_change_id)  = Trim(Request("txthToOrgChangeId")) 'gChangeOrgId    

    I6_a_asset_chg(A507_I6_chg_fg)        = "05"				       '변동구분	
	I6_a_asset_chg(A507_I6_temp_gl_no)		= Trim(Request("txtFrTempGLNo"))				 
	I6_a_asset_chg(A507_I6_to_gl_no)		= Trim(Request("txtToTempGLNo"))				 
	I6_a_asset_chg(A507_I6_gl_no)			= Trim(Request("txtFrGlNo"))			
	I6_a_asset_chg(A507_I6_to_gl_no)			= Trim(Request("txtToTempGLNo"))			
	I6_a_asset_chg(A507_I6_asset_chg_desc) = trim (Request("txtDesc"))            '변동사유       

    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
		I8_ief_supplied_select_char = "C"
		
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
		I8_ief_supplied_select_char = "U"

    End If

	'-----------------------
	'Com Action Area
	'-----------------------
	
	 E2_a_asset_chg = iPAAG035.AS0031_MANAGE_SVR( gStrGloBalCollection ,iCommandSent , I2_b_acct_dept ,I3_b_biz_partner_bp_cd , _
	                                             I4_b_currency_currency ,I5_b_acct_dept ,I6_a_asset_chg ,Request("txtSpread"),I7_a_asset_master_asst_no , _
	                                             I8_ief_supplied_select_char ,I9_a_data_auth) 

 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG035 =nothing	
       Exit Sub
    End If 

    Set iPAAG035 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtChgNo.Value=  """ & ConvSPChars(E2_a_asset_chg) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 															'☜: 조화가 성공 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  
	Response.End
	
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

	redim I6_a_asset_chg (A509_E7_vat_loc_amt)

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
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
	On Error Resume Next														'☜: 
    Err.Clear                                                               '☜: Protect system from crashing

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


'    vAs0031.CommandSent = "DELETE"
    I6_a_asset_chg(A507_I6_chg_no)   = Trim(Request("txtChgNo"))   
    I6_a_asset_chg(A507_I6_chg_fg)   = "05"

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
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
	
%>
