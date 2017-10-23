<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb1
'*  4. Program Name         : �����ڻ� MASTER ����
'*  5. Program Desc         : �����ڻ꺰 MASTER�� ����, ��ȸ
'*  6. Comproxy List        : +As0041ManageSvr
'                             +As0049LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/29
'*  8. Modified date(Last)  : 2000/09/14
'*  9. Modifier (First)     : ���ͼ�
'* 10. Modifier (Last)      : hersheys
'* 11. Comment              :
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd       
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")                                                        '��: Hide Processing message

	gChangeOrgId = GetGlobalInf("gChangeOrgId")

	Dim lgOpModeCRUD
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

	Response.End 
    'Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	'------------------
	' ����, ��� ���� 
	'------------------
	Dim iPAAG015
	Dim I1_a_asset_master_asst_no
	Dim I2_b_acct_dept_org_change_id
	Dim E1_a_batch							
	Dim E2_a_asset_master
	Dim E3_a_mcs
	Dim E4_a_acct		
	Dim E5_a_asset_acct
	Dim E6_a_asset_depr_rate	
	Dim E7_a_asset_depr_rate	
	Dim E8_a_mcs
	Dim E9_b_biz_partner
	Dim E10_b_acct_dept
	Dim E11_a_asset_acq
	Dim E12_b_cost_center	'�ڽ�Ʈ����(��ȯ) : air

    'export a_batch						[��ǥ�ϰ�����]
    Const A512_E1_gl_dt = 0		'ȸ������
    
    'export a_asset_master				[�ڻ긶����]
    Const A512_E2_asst_no = 0			'�ڻ��ȣ
    Const A512_E2_asst_nm = 1			'�ڻ��
    Const A512_E2_ref_no = 2			'ǰ��׷�
    Const A512_E2_reg_dt = 3			'�ڻ�������
    Const A512_E2_acq_amt = 4			'���ݾ�
    Const A512_E2_doc_cur = 5			'�ŷ���ȭ
    Const A512_E2_xch_rate = 6			'ȯ��
    Const A512_E2_acq_loc_amt = 7		'���ݾ�(�ڱ�)
    Const A512_E2_acq_qty = 8			'������
    Const A512_E2_inv_qty = 9			'������
    Const A512_E2_tax_dur_yrs = 10		'�������س�����
    Const A512_E2_cas_dur_yrs = 11		'���ȸ����س�����
    Const A512_E2_tax_end_l_term_cpt_tot_amt = 12	'�����������⸻�ں�������ݾ�
    Const A512_E2_cas_end_l_term_cpt_tot_amt = 13	'���ȸ��������⸻�ں�������ݾ�
    Const A512_E2_tax_end_l_term_depr_tot_amt = 14	'�����������⸻�����󰢴���ݾ�
    Const A512_E2_cas_end_l_term_depr_tot_amt = 15  '���ȸ��������⸻�����󰢴���ݾ�
    Const A512_E2_tax_end_l_term_bal_amt = 16		'�������ع̻��ܾ�
    Const A512_E2_cas_end_l_term_bal_amt = 17		'���ȸ����ع̻��ܾ�
    Const A512_E2_tax_depr_end_yyyymm = 18			'�������ػ󰢿Ϸ���
    Const A512_E2_cas_depr_end_yyyymm = 19			'���ȸ����ػ󰢿Ϸ���
    Const A512_E2_tax_depr_sts = 20		'�������ػ󰢻���
    Const A512_E2_cas_depr_sts = 21		'���ȸ����ػ󰢻���
    Const A512_E2_spec = 22				'�뵵/�԰�
    Const A512_E2_asset_desc = 23		'����
    Const A512_E2_start_depr_yymm = 24	'�����󰢽��۳��
    Const A512_E2_gl_no = 25			'��ǥ��ȣ
    Const A512_E2_temp_gl_no = 26		'������ǥ��ȣ
    Const A512_E2_temp_fg1 = 27			'������ǥ��
    Const A512_E2_disuse_fg = 28		'�Ű�/���Ϸ�
    Const A512_E2_disuse_yymm = 29		'�Ű�/���Ϸ���
    Const A512_E2_vat_rate = 30			'�ΰ�����
    Const A512_E2_net_amt = 31			'���ް���
    Const A512_E2_net_loc_amt = 32		'���ް���(�ڱ�)
    Const A512_E2_tax_dur_mnth = 33		'�������س������
    Const A512_E2_cas_dur_mnth = 34		'���ȸ����ر��س������
	
	'export_start_yymm a_mcs		[a_mcs]
	Const A512_E3_txt_from_dt = 0   'txt_from_dt
	
    'export a_acct					[�����ڵ�]
    Const A512_E4_acct_cd = 0		'�����ڵ�    
    Const A512_E4_acct_nm = 1		'�����ܸ�

    'export a_asset_acct			[�ڻ�����ڵ�]
    Const A512_E5_depr_mthd = 0		'�󰢹��
    Const A512_E5_dur_yrs = 1		'������

    'export_tax a_asset_depr_rate   [�����󰢷�����]
    Const A512_E6_depr_rate = 0     '�����󰢷�(��������)

    'export_cas a_asset_depr_rate	[�����󰢷�����]
    Const A512_E7_depr_rate = 0     '�����󰢷�(���ȸ�����)

    'export a_mcs
    Const A512_E8_txt_from_dt = 0    
    Const A512_E8_txt_to_dt = 1

    'export b_biz_partner			[�ŷ�ó]
    Const A512_E9_bp_cd = 0			'�ŷ�ó�ڵ�          
    Const A512_E9_bp_nm = 1			'�ŷ�ó(���)

    'export b_acct_dept				[ȸ��μ�����]
    Const A512_E10_dept_cd = 0		'�μ��ڵ�
    Const A512_E10_dept_nm = 1		'�μ����

    'export a_asset_acq				[�ڻ����]
    Const A512_E11_acq_fg = 0       '��汸��
    Const A512_E11_acq_no = 1		'�ڻ�����ȣ
    Const A512_E11_gl_no = 2		'��ǥ��ȣ
    Const A512_E11_temp_gl_no = 3	'������ǥ��ȣ

    'export B_COST_CENTER			[�ڽ�Ʈ����]
    Const A512_E12_cost_cd = 0       'cost code
    Const A512_E12_cost_nm = 1       'cost name


	' -- ���Ѱ����߰�
	Const A512_I3_a_data_auth_data_BizAreaCd = 0
	Const A512_I3_a_data_auth_data_internal_cd = 1
	Const A512_I3_a_data_auth_data_sub_internal_cd = 2
	Const A512_I3_a_data_auth_data_auth_usr_id = 3

	Dim I3_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ����

  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A512_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	'------------------
	' Data Matching
	'------------------
	'ReDim E2_a_asset_master(34)
	'ReDim E4_a_acct(1)
	'ReDim E5_a_asset_acct(1)
	'ReDim E8_a_mcs(1)
	'ReDim E9_b_biz_partner(1)
	'ReDim E10_b_acct_dept(1)
	'ReDim E11_a_asset_acq(3)
	
	'*** Import Data ***
	I1_a_asset_master_asst_no = Request("txtCondAsstNo")	'�ڻ��ȣ
	I2_b_acct_dept_org_change_id = gChangeOrgId				'��������ID

	'------------------
	' ��û ���� ó�� 
	'------------------
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Set iPAAG015 = Server.CreateObject("PAAG015_KO441.cAAS0049LkUpSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

	Call iPAAG015.AS0049_LOOKUP_SVR(gStrGloBalCollection, _ 
									I1_a_asset_master_asst_no, _
									I2_b_acct_dept_org_change_id, _
									E1_a_batch, _
									E2_a_asset_master, _
									E3_a_mcs, _
									E4_a_acct, _
									E5_a_asset_acct, _
									E6_a_asset_depr_rate, _
									E7_a_asset_depr_rate, _
									E8_a_mcs, _
									E9_b_biz_partner, _
									E10_b_acct_dept, _
									E11_a_asset_acq, _
									E12_b_cost_center, _
									I3_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG015 = Nothing
       'Response.End
       Exit Sub
    End If    
    Set iPAAG015 = Nothing

	'---------------------
	' HTML �� ��� ������ 
	'---------------------
	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

    '*** �⺻���� ***
    Response.Write "	.frm1.txtAsstNm.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_asst_nm))		& """" & vbCr '�ڻ��
    Response.Write "	.frm1.txtRefNo.value     = """ & ConvSPChars(E2_a_asset_master(A512_E2_ref_no))			& """" & vbCr 'ǰ��׷�            
    Response.Write "	.frm1.txtDeptCd.value    = """ & ConvSPChars(E10_b_acct_dept(A512_E10_dept_cd))			& """" & vbCr '�����μ��ڵ�            
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E10_b_acct_dept(A512_E10_dept_nm))			& """" & vbCr '�����μ���
    
    Response.Write "	.frm1.txtCostCd.value    = """ & ConvSPChars(E12_b_cost_center(A512_E12_cost_cd))			& """" & vbCr '�ڽ�Ʈ�����ڵ�            
    Response.Write "	.frm1.txtCostNm.value    = """ & ConvSPChars(E12_b_cost_center(A512_E12_cost_nm))			& """" & vbCr '�ڽ�Ʈ���͸�
    
    Response.Write "	.frm1.txtRegDt.text     = """ & UNIDateClientFormat(E2_a_asset_master(A512_E2_reg_dt))	& """" & vbCr '�������(�ڻ�������)
    Response.Write "	.frm1.txtDocCur.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_doc_cur))		& """" & vbCr '�ŷ���ȭ
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_xch_rate),   ggExchRate.DecPoint, 0)		& """" & vbCr 'ȯ��            
    Response.Write "	.frm1.txtAcqAmt.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_amt),    ggAmtOfMoney.DecPoint, 0)	& """" & vbCr '���ݾ�            
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_loc_amt), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr '���ݾ�(�ڱ�)            
    Response.Write "	.frm1.txtAcqQty.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_qty),    ggQty.DecPoint, 0)			& """" & vbCr '������            
    Response.Write "	.frm1.txtInvQty.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_inv_qty),    ggQty.DecPoint, 0)			& """" & vbCr '������            
    Response.Write "	.frm1.txtAcctCd.value    = """ & ConvSPChars(E4_a_acct(A512_E4_acct_cd))				& """" & vbCr '�����ڵ�            
    Response.Write "	.frm1.txtAcctNm.value    = """ & ConvSPChars(E4_a_acct(A512_E4_acct_nm))				& """" & vbCr '�����ڵ��            
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E9_b_biz_partner(A512_E9_bp_cd))			& """" & vbCr '���԰ŷ�ó�ڵ�            
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E9_b_biz_partner(A512_E9_bp_nm))			& """" & vbCr '���԰ŷ�ó��            
    Response.Write "	.frm1.cboAcqFg.value     = """ & E11_a_asset_acq(A512_E11_acq_fg)						& """" & vbCr '��汸��            
    Response.Write "	.frm1.txtSpec.value      = """ & Trim(ConvSPChars(E2_a_asset_master(A512_E2_spec)))			& """" & vbCr '����/�뵵/ũ��            
    Response.Write "	.frm1.txtDesc.value      = """ & Trim(ConvSPChars(E2_a_asset_master(A512_E2_asset_desc)))	& """" & vbCr '����            
    Response.Write "	.frm1.txtDeprFrdt.text  = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_start_depr_yymm))						& """" & vbCr '�����󰢽��۳��            
    
    '*** ���⸻ �󰢳���:��������(�ڱ�) ***
	'Response.Write "	.frm1.txtTaxDurYrs.value     = """ & E2_a_asset_master(A512_E2_tax_dur_yrs)				& """" & vbCr '���뿬��            
	Response.Write "	.frm1.txtTaxDurYrs.value     = """ & E2_a_asset_master(A512_E2_tax_dur_mnth)			& """" & vbCr '�������	>>air            
	Response.Write "	.frm1.txtTaxDeprRate.value   = """ & UNINumClientFormat(E6_a_asset_depr_rate(A512_E6_depr_rate), ggExchRate.DecPoint, 0)				  & """" & vbCr '����            
	Response.Write "	.frm1.txtTaxDeprTotAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_depr_tot_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr '�󰢴���
	Response.Write "	.frm1.txtTaxCptTotAmt.value  = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_cpt_tot_amt), ggAmtOfMoney.DecPoint, 0)  & """" & vbCr '�ں������⴩��            
	Response.Write "	.frm1.txtTaxBalAmt.value     = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_bal_amt), ggAmtOfMoney.DecPoint, 0)	  & """" & vbCr '�̻��ܾ�            
	Response.Write "	.frm1.cboTaxDeprSts.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_tax_depr_sts))								& """" & vbCr '�󰢻���            
	Response.Write "	.frm1.txtTaxDeprEnd.text    = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_tax_depr_end_yyyymm))				& """" & vbCr '�󰢿Ϸ���            
	
    '*** ���⸻ �󰢳���:���ȸ�����(�ڱ�) ***
	'Response.Write "	.frm1.txtCasDurYrs.value     = """ & E2_a_asset_master(A512_E2_cas_dur_yrs)				& """" & vbCr '���뿬��            
	Response.Write "	.frm1.txtCasDurYrs.value     = """ & E2_a_asset_master(A512_E2_cas_dur_mnth)			& """" & vbCr '������� >>air           
	Response.Write "	.frm1.txtCasDeprRate.value   = """ & UNINumClientFormat(E7_a_asset_depr_rate(A512_E7_depr_rate), ggExchRate.DecPoint, 0)				  & """" & vbCr '����            
	Response.Write "	.frm1.txtCasDeprTotAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_depr_tot_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr '�󰢴���
	Response.Write "	.frm1.txtCasCptTotAmt.value  = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_cpt_tot_amt), ggAmtOfMoney.DecPoint, 0)  & """" & vbCr '�ں������⴩��            
	Response.Write "	.frm1.txtCasBalAmt.value     = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_bal_amt), ggAmtOfMoney.DecPoint, 0)	  & """" & vbCr '�̻��ܾ�            
	Response.Write "	.frm1.cboCasDeprSts.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_cas_depr_sts))								& """" & vbCr '�󰢻���            
	Response.Write "	.frm1.txtCasDeprEnd.text    = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_cas_depr_end_yyyymm))				& """" & vbCr '�󰢿Ϸ���            
    
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    'Response.End

End Sub	
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
'Call ServerMesgBox("a", vbInformation, I_MKSCRIPT)	
	'------------------
	' ����, ��� ���� 
	'------------------
    Dim iPAAG015
	Dim I1_a_asset_master
	
    'import a_asset_master
    Const A510_I1_asst_no = 0						'�ڻ��ȣ
    Const A510_I1_ref_no = 1						'ǰ��׷�
    Const A510_I1_tax_dur_yrs = 2					'�������س�����
    Const A510_I1_cas_dur_yrs = 3					'���ȸ����س�����
    Const A510_I1_tax_end_l_term_cpt_tot_amt = 4	'�����������⸻�ں�������ݾ�
    Const A510_I1_cas_end_l_term_cpt_tot_amt = 5	'���ȸ��������⸻�ں�������ݾ�
    Const A510_I1_tax_end_l_term_depr_tot_amt = 6	'�����������⸻�����󰢴���ݾ�
    Const A510_I1_cas_end_l_term_depr_tot_amt = 7	'���ȸ��������⸻�����󰢴���ݾ�
    Const A510_I1_tax_end_l_term_bal_amt = 8		'�������ع̻��ܾ�
    Const A510_I1_cas_end_l_term_bal_amt = 9		'���ȸ����ع̻��ܾ�
    Const A510_I1_tax_depr_sts = 10					'�������ػ󰢻���
    Const A510_I1_cas_depr_sts = 11					'���ȸ����ػ󰢻���
    Const A510_I1_tax_depr_end_yyyymm = 12			'�������ػ󰢿Ϸ���
    Const A510_I1_cas_depr_end_yyyymm = 13			'���ȸ����ػ󰢿Ϸ���
    Const A510_I1_updt_user_id = 14					'User ID
    Const A510_I1_start_depr_yymm = 15				'�����󰢽��۳��
    Const A510_I1_vat_rate = 16						'�ΰ�����
    Const A510_I1_net_amt = 17						'���ް���
    Const A510_I1_net_loc_amt = 18					'���ް���(�ڱ�)
    Const A510_I1_temp_fg1 = 19						'������ǥ��
    Const A510_I1_asst_nm = 20						'�ڻ��
    Const A510_I1_spec = 21							'�뵵/�԰�
    Const A510_I1_asset_desc = 22					'����
    Const A510_I1_tax_dur_mnth = 23					'�������س������
    Const A510_I1_cas_dur_mnth = 24					'���ȸ����ر��س������
    Const A510_I1_COST_CD = 25						'�ڽ�Ʈ���� >>AIR

	' -- ���Ѱ����߰�
	Const A512_I2_a_data_auth_data_BizAreaCd = 0
	Const A512_I2_a_data_auth_data_internal_cd = 1
	Const A512_I2_a_data_auth_data_sub_internal_cd = 2
	Const A512_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ����

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A512_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	'------------------
	' Data Matching
	'------------------
	'*** Import Data ***
	Redim I1_a_asset_master(25)
	
	I1_a_asset_master(A510_I1_asst_no)						= Trim(Request("txtCondAsstNo")) 			   '�ڻ��ȣ                    
    I1_a_asset_master(A510_I1_ref_no)						= Trim(Request("txtRefNo"))					   'ǰ��׷�                    
    I1_a_asset_master(A510_I1_tax_dur_yrs)					= UNIConvNum(Request("txtTaxDurYrs"), 0)       '�������س�����            
    I1_a_asset_master(A510_I1_cas_dur_yrs)					= UNIConvNum(Request("txtCasDurYrs"), 0)       '���ȸ����س�����        
    I1_a_asset_master(A510_I1_tax_end_l_term_cpt_tot_amt)	= UNIConvNum(Request("txtTaxCptTotAmt"), 0)   '�����������⸻�ں�������ݾ�
    I1_a_asset_master(A510_I1_cas_end_l_term_cpt_tot_amt)	= UNIConvNum(Request("txtCasCptTotAmt"), 0)   '���ȸ��������⸻�ں�������
    I1_a_asset_master(A510_I1_tax_end_l_term_depr_tot_amt)	= UNIConvNum(Request("txtTaxDeprTotAmt"), 0)   '�����������⸻�����󰢴����
    I1_a_asset_master(A510_I1_cas_end_l_term_depr_tot_amt)	= UNIConvNum(Request("txtCasDeprTotAmt"), 0)   '���ȸ��������⸻�����󰢴�
    I1_a_asset_master(A510_I1_tax_end_l_term_bal_amt)		= UNIConvNum(Request("txtTaxBalAmt"), 0)       '�������ع̻��ܾ�          
    I1_a_asset_master(A510_I1_cas_end_l_term_bal_amt)		= UNIConvNum(Request("txtCasBalAmt"), 0)       '���ȸ����ع̻��ܾ�      
    I1_a_asset_master(A510_I1_tax_depr_sts)					= Request("cboTaxDeprSts")      '�������ػ󰢻���            
    I1_a_asset_master(A510_I1_cas_depr_sts)					= Request("cboCasDeprSts")      '���ȸ����ػ󰢻���        
    I1_a_asset_master(A510_I1_tax_depr_end_yyyymm)			= Request("txtTaxDeprEnd")      '�������ػ󰢿Ϸ���        
    I1_a_asset_master(A510_I1_cas_depr_end_yyyymm)			= Request("txtCasDeprEnd")      '���ȸ����ػ󰢿Ϸ���    
    I1_a_asset_master(A510_I1_updt_user_id)					= gUsrID						'User ID                     
    I1_a_asset_master(A510_I1_start_depr_yymm)				= Request("txtDeprFrdt")        '�����󰢽��۳��            
    I1_a_asset_master(A510_I1_vat_rate)						= 0								'? �ΰ�����                    
    I1_a_asset_master(A510_I1_net_amt)						= 0								'? ���ް���                    
    I1_a_asset_master(A510_I1_net_loc_amt)					= 0								'? ���ް���(�ڱ�)              
   'I1_a_asset_master(A510_I1_temp_fg1)						= ?						  	    '������ǥ��                  
    I1_a_asset_master(A510_I1_asst_nm)						= Request("txtAsstNm")          '�ڻ��                      
    I1_a_asset_master(A510_I1_spec)							= Request("txtSpec")            '�뵵/�԰�                   
    I1_a_asset_master(A510_I1_asset_desc)					= Request("txtDesc")            '����                        
   'I1_a_asset_master(A510_I1_tax_dur_mnth)					= ?                             '�������س������            
   'I1_a_asset_master(A510_I1_cas_dur_mnth)					= ?                             '���ȸ����ر��س������    
	I1_a_asset_master(A510_I1_COST_CD)						= Trim(Request("txtCostCd"))	'�ڽ�Ʈ���� >>AIR
	'------------------
'Call ServerMesgBox(Request("txtCostCd"), vbInformation, I_MKSCRIPT)	
	' ��û ���� ó�� 
	'------------------
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    Set iPAAG015 = Server.CreateObject("PAAG015_KO441.cAAS0041MngSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    
    Call iPAAG015.AS0041_MANAGE_SVR(gStrGloBalCollection, I1_a_asset_master, I2_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG015 = Nothing
       'Response.End
       Exit Sub
    End If    
    
    Set iPAAG015 = Nothing
	
	'---------------------
	' HTML �� ��� ������ 
	'---------------------
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	

%>
