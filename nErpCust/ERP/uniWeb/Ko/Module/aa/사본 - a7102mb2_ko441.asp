<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : �����ڻ����
'*  3. Program ID           : a7102mb2
'*  4. Program Name         : �����ڻ���泻�����
'*  5. Program Desc         : �����ڻ꺰 ��泻���� ���,����
'*  6. Comproxy List        : +As0021ManageSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2001/05/24
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : ������
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ�
On Error Resume Next														'��: 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide


    Call LoadBasisGlobalInf()
	'Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	'Call LoadBNumericFormat("I", "*")
'	Dim gChangeOrgId 'iChangeOrgId
	gChangeOrgId = request("hOrgChangeId")
	
	
'-------------------------
' ����, ��� ���� 
'-------------------------
	Dim iPAAG010																'�� : ����� ComProxy Dll ��� ����
	Dim lgIntFlgMode

	lgIntFlgMode = CInt(Request("txtFlgMode"))									'��: ����� Create/Update �Ǻ�
		
	'Dim IntRows
	'Dim IntCols
	'Dim vbIntRet
	'Dim lEndRow
	'Dim boolCheck
	'Dim lgIntFlgMode
	'Dim LngMaxRow_m
	'Dim LngMaxRow_i

    '[Import ����]
	Dim I1_a_acct_trans_type	'�ŷ�����
	Dim I2_b_currency			'�ŷ���ȭ
	Dim I3_a_asset_acq			'Control Data
	Dim I4_ief_supplied			'����Ű(C,U)
	Dim IG1_import_mst_grp		'Master Data
	Dim IG2_import_itm_grp		'��ݳ���
	Dim I5_b_biz_partner		'�ŷ�ó
	Dim I6_b_acct_dept			'�μ�����
	Dim I7_a_asset_acq			'ä����������(�����ޱ�)
	Dim I8_a_batch				'��ǥ����
	Dim E3_a_asset_acq

	'[Import ���]
    Const A504_I1_trans_type = 0	'import a_acct_trans_type

    Const A504_I2_currency = 0		'import b_currency

    Const A504_I3_acq_no = 0		'import a_asset_acq
    Const A504_I3_acq_dt = 1
    Const A504_I3_acq_fg = 2
    Const A504_I3_doc_cur = 3
    Const A504_I3_xch_rate = 4
    Const A504_I3_tot_acq_amt = 5
    Const A504_I3_tot_acq_loc_amt = 6
    Const A504_I3_extra_acq_amt = 7
    Const A504_I3_extra_acq_loc_amt = 8
    Const A504_I3_vat_type = 9
    Const A504_I3_vat_make_fg = 10
    Const A504_I3_vat_amt = 11
    Const A504_I3_vat_loc_amt = 12
    Const A504_I3_ref_no = 13
    '20030301	�����ޱݰ��� �߰�
    Const A504_I3_ap_acct_cd = 14
    Const A504_I3_ap_due_dt = 15
    Const A504_I3_ap_amt = 16
    Const A504_I3_ap_loc_amt = 17
    Const A504_I3_acq_desc = 18
    Const A504_I3_ap_no = 19
    Const A504_I3_gl_no = 20
    Const A504_I3_temp_gl_no = 21
    Const A504_I3_internal_cd = 22
    Const A504_I3_insrt_user_id = 23
    Const A504_I3_insrt_dt = 24
    Const A504_I3_updt_user_id = 25
    Const A504_I3_updt_dt = 26
    Const A504_I3_vat_io_fg = 27
    Const A504_I3_vat_rate = 28
    Const A504_I3_issued_dt = 29
    Const A504_I3_tax_biz_area_cd = 30
    Const A504_I3_Credit_card_No = 31															'�ſ�ī�� ��ȣ �߰�
   

    Const A504_I4_select_char = 0	'import_mode_fg ief_supplied

    Const A504_I5_bp_cd = 0			'import b_biz_partner

    Const A504_I6_org_change_id = 0	'import b_acct_dept
    Const A504_I6_dept_cd = 1

    Const A504_I7_ap_due_dt = 0		'import_null_dt a_asset_acq

    Const A504_I8_gl_dt = 0			'import_a_batch a_batch

    '[IMPORTS Group ���]	'20080305 �ּ�ó�� air
    'Group Name : import_mst_grp
    'Const A504_IG1_I1_acct_cd = 0        'View Name : import_mst_itm a_acct
    'Const A504_IG1_I2_select_char = 1    'View Name : import_mst_itm ief_supplied
    'Const A504_IG1_I3_org_change_id = 2  'View Name : import_mst_itm b_acct_dept
    'Const A504_IG1_I3_dept_cd = 3
    'Const A504_IG1_I4_asst_no = 4        'View Name : import_mst_itm a_asset_master
    'Const A504_IG1_I4_asst_nm = 5
    'Const A504_IG1_I4_reg_dt = 6
    'Const A504_IG1_I4_spec = 7
    'Const A504_IG1_I4_doc_cur = 8
    'Const A504_IG1_I4_xch_rate = 9
    'Const A504_IG1_I4_acq_amt = 10
    'Const A504_IG1_I4_acq_loc_amt = 11
    'Const A504_IG1_I4_acq_qty = 12
    'Const A504_IG1_I4_inv_qty = 13
    'Const A504_IG1_I4_asset_desc = 14
    'Const A504_IG1_I4_ref_no = 15
    'Const A504_IG1_I4_tax_dur_yrs = 16
    'Const A504_IG1_I4_cas_dur_yrs = 17
    'Const A504_IG1_I4_tax_end_l_term_cpt_tot_amt = 18
    'Const A504_IG1_I4_cas_end_l_term_cpt_tot_amt = 19
    'Const A504_IG1_I4_tax_end_l_term_depr_tot_amt = 20
    'Const A504_IG1_I4_cas_end_l_term_depr_tot_amt = 21
    'Const A504_IG1_I4_tax_end_l_term_bal_amt = 22
    'Const A504_IG1_I4_cas_end_l_term_bal_amt = 23
    'Const A504_IG1_I4_tax_depr_sts = 24
    'Const A504_IG1_I4_cas_depr_sts = 25
    'Const A504_IG1_I4_tax_depr_end_yyyymm = 26
    'Const A504_IG1_I4_cas_depr_end_yyyymm = 27
    'Const A504_IG1_I4_start_depr_yymm = 28
    'Const A504_IG1_I4_tax_dur_mnth = 29
    'Const A504_IG1_I4_cas_dur_mnth = 30
    '
    ''Group Name : import_itm_grp
    'Const A504_IG2_I1_select_char = 0    'View Name : import_itm_item ief_supplied
    'Const A504_IG2_I2_acq_seq = 1        'View Name : import_itm_item a_asset_acq_item
    'Const A504_IG2_I2_paym_type = 2
    'Const A504_IG2_I2_paym_amt = 3
    'Const A504_IG2_I2_paym_loc_amt = 4
    'Const A504_IG2_I2_note_no = 5
    'Const A504_IG2_I3_bank_acct_no = 6   'View Name : import_itm_item b_bank_acct
	
	'Export
	Const A073_E3_acq_no = 0             'View Name : export a_asset_acq
'-------------------------   
' Data Matching
'------------------------- 
	Redim I1_a_acct_trans_type(0)
	'Redim I2_b_currency(0)
	'20030301	�����ޱݰ��� �߰�	
	Redim I3_a_asset_acq(31)
	Redim I4_ief_supplied(0)
	Redim I5_b_biz_partner(0)
	Redim I6_b_acct_dept(1)	
	Redim I7_a_asset_acq(0)
	Redim I8_a_batch(0)
	Redim E3_a_asset_acq(0)


	' -- ���Ѱ����߰�
	Const A504_I9_a_data_auth_data_BizAreaCd = 0
	Const A504_I9_a_data_auth_data_internal_cd = 1
	Const A504_I9_a_data_auth_data_sub_internal_cd = 2
	Const A504_I9_a_data_auth_data_auth_usr_id = 3

	Dim I9_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ����

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A504_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	
	I1_a_acct_trans_type(A504_I1_trans_type)  = "AS001"		'�ŷ�����
	
	I2_b_currency = gCurrency	'�ŷ���ȭ	'(A504_I2_currency)

	'Control Data
    I3_a_asset_acq(A504_I3_acq_no)			  = Request("txtAcqNo")						'00 ����ȣ
    I3_a_asset_acq(A504_I3_acq_dt)			  = UNIConvDate(Request("txtAcqDt"))		'01 �������
    I3_a_asset_acq(A504_I3_acq_fg)			  = Request("cboAcqFg")						'02 ��汸��
    I3_a_asset_acq(A504_I3_doc_cur)			  = UCase(Trim(Request("txtDocCur")))				'03 �ŷ���ȭ
    I3_a_asset_acq(A504_I3_xch_rate)		  = UNIConvNum(Request("txtXchRate"),0)		'04 ȯ��
    I3_a_asset_acq(A504_I3_tot_acq_amt)       = UNIConvNum(Request("txtAcqAmt"),0)		'05 �����ݾ�
    I3_a_asset_acq(A504_I3_tot_acq_loc_amt)   = UNIConvNum(Request("txtAcqLocAmt"),0)	'06 �����ݾ�(�ڱ�)
   'I3_a_asset_acq(A504_I3_extra_acq_amt)     											'07 �δ���
   'I3_a_asset_acq(A504_I3_extra_acq_loc_amt) 											'08 �δ���(�ڱ�)
    I3_a_asset_acq(A504_I3_vat_type)		  = UCase(Request("txtVatType"))					'09 �ΰ�������
   'I3_a_asset_acq(A504_I3_vat_make_fg)													'10 �ΰ��� ��������
    I3_a_asset_acq(A504_I3_vat_amt)			  = UNIConvNum(Request("txtVatAmt"),0)		'11 �ΰ����ݾ�
    I3_a_asset_acq(A504_I3_vat_loc_amt)		  = UNIConvNum(Request("txtVatLocAmt"),0)	'12 �ΰ����ݾ�(�ڱ�)
   'I3_a_asset_acq(A504_I3_ref_no)														'13 ������ȣ(Master Spread)
   '20030301	�����ޱݰ��� �߰�
    I3_a_asset_acq(A504_I3_ap_acct_cd)		  = Trim(UCase(Request("txtApAcctCd")))		'14 �����ޱ� ����
    I3_a_asset_acq(A504_I3_ap_due_dt)		  = UNIConvDate(Request("txtApDueDt"))		'14 �����ޱ� ��������
    I3_a_asset_acq(A504_I3_ap_amt)			  = UNIConvNum(Request("txtApAmt"),0)		'15 �����ޱݾ�
    I3_a_asset_acq(A504_I3_ap_loc_amt)		  = UNIConvNum(Request("txtApLocAmt"),0)	'16 �����ޱݾ�(�ڱ�)
    I3_a_asset_acq(A504_I3_acq_desc)		  = Trim(Request("txtDesc"))					'17 ����(Master Spread)
    I3_a_asset_acq(A504_I3_ap_no)			  = Trim(Request("txtApNo"))				'18 �����ޱ� ��ȣ
    I3_a_asset_acq(A504_I3_gl_no)			  = Trim(Request("txtGLNo"))				'19 ȸ����ǥ��ȣ
    I3_a_asset_acq(A504_I3_temp_gl_no)		  = Trim(Request("txtTempGLNo"))			'20 ������ǥ��ȣ
   'I3_a_asset_acq(A504_I3_internal_cd)		  											'21 ���κμ��ڵ�
   'I3_a_asset_acq(A504_I3_insrt_user_id)	  											'22 �Է���
   'I3_a_asset_acq(A504_I3_insrt_dt)		  											'23 �Է���
   'I3_a_asset_acq(A504_I3_updt_user_id)	  											'24 ������
   'I3_a_asset_acq(A504_I3_updt_dt)			  											'25 ������
    I3_a_asset_acq(A504_I3_vat_io_fg)		 = "I" 											'26 �ΰ��� ����/���� ����
    I3_a_asset_acq(A504_I3_vat_rate)		  = UNIConvNum(Request("txtVatRate"),0)		'28 �ΰ�����
	
	If Trim(Request("txtIssuedDt")) <>"" Then
		I3_a_asset_acq(A504_I3_issued_dt)	=	UNIConvDate(Request("txtIssuedDt"))     ' 10�� ���� ��ġ �߰�
	End If																				'29 ��꼭 ������
	
	I3_a_asset_acq(A504_I3_tax_biz_area_cd)	=	Trim(Request("txtReportAreaCd"))		'30 �Ű�����
	I3_a_asset_acq(A504_I3_Credit_card_No)	=	Trim(Request("txtCardNo"))										'31 �ſ�ī�� ��ȣ (����)

    '����Ű(C,U)
	If lgIntFlgMode = OPMD_CMODE Then
		I4_ief_supplied(A504_I4_select_char) = "C"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		I4_ief_supplied(A504_I4_select_char) = "U"
	End If
    
	IG1_import_mst_grp = Request("txtSpread_m")		'���󼼳��� Spread
	IG2_import_itm_grp = Request("txtSpread_i")		'��ݳ��� Spread

	I5_b_biz_partner(A504_I5_bp_cd) = Trim(Request("txtBpCd"))	'�ŷ�ó
	
	'�μ�����
	I6_b_acct_dept(A504_I6_org_change_id) = gChangeOrgId
	I6_b_acct_dept(A504_I6_dept_cd) = Trim(Request("txtDeptCd"))	'���μ�
	
	'ä����������(�����ޱ�)
	I7_a_asset_acq(A504_I7_ap_due_dt) = UNIConvDate(Request("txtApDueDt"))
	
	'��ǥ����
	I8_a_batch(A504_I8_gl_dt) = UNIConvDate(Request("txtGLDt"))				

'-------------------------   
' ���� ó�� 
'-------------------------    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 

    Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqMngSvr")	'AIR :: PAAG010 -> PAAG010_KO441


    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If  
    

	call iPAAG010.AS0021_ACQ_MANAGE_SVR(gStrGlobalCollection, _
										I1_a_acct_trans_type, _
										I2_b_currency, _
										I3_a_asset_acq, _							
										I4_ief_supplied, _
										IG1_import_mst_grp, _
										IG2_import_itm_grp, _
										I5_b_biz_partner, _	
										I6_b_acct_dept, _
										I7_a_asset_acq, _		
										I8_a_batch, _
										E1_a_asset_master, _
										E3_a_asset_acq, _
										I9_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing
    
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write "	parent.frm1.txtAcqNo.value = """ & E3_a_asset_acq(A073_E3_acq_no) & """" & vbCr '�ڻ�����ȣ        
    Response.Write "	parent.DbSaveOk()												   	   " & vbCr
    Response.Write " </Script>					" & vbCr
%>
