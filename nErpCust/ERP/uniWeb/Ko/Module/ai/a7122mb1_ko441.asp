<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<% 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb1
'*  4. Program Name         : �����ڻ���泻����� 
'*  5. Program Desc         : �����ڻ���泻���� ��ȸ 
'*  6. Comproxy List        : +As0029LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : ������ 
'* 10. Modifier (Last)      : ������ 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'===================================================================
'						1. Include
'===================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'===================================================================
'						2. ���Ǻ� 
'===================================================================
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
On Error Resume Next														'��: 

    Call LoadBasisGlobalInf()
	Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'-------------------------  
' 2.1 ���� üũ 
'-------------------------  
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then											'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)		'��: ��ȸ �����ε� �ٸ� ���·� ��û�� ���� ���, �ʿ������ ���� ��, �޼����� ID������ ����ؾ� �� 
	Response.End 
ElseIf Trim(Request("txtAcqNo")) = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)						'��:
	Response.End 
End If

'===================================================================
'						2. ���� ó�� ����� 
'===================================================================
'-------------------------  
' 2.1. ����, ��� ���� 
'-------------------------  

    Dim iPAAG010
    
    'Import ���� 
    Dim I1_ief_supplied		'������ 
    Dim I2_a_asset_acq_item	'�ڻ������� 
    Dim I3_a_asset_master	'�ڻ��ȣ 
    Dim I4_a_asset_acq		'�ڻ�����ȣ 

    'Export ���� 
    Dim E1_b_minor			'�ΰ��������� 
    Dim E2_a_batch			'��ǥ���� 
    Dim E3_a_asset_acq_item	'�ڻ������� 
    Dim E4_a_asset_master	'�ڻ��ȣ 
    Dim EG1_exp_group
    Dim E5_b_biz_partner
    Dim E6_b_acct_dept
    Dim E7_a_asset_acq
    Dim EG2_export_itm_grp
     
    Dim iLngRow,iLngCol
    Dim idate, StrDate
    Dim iStrData
    Dim iStrData2
    Dim iStsNm	'�󰢻��¸� 
    Dim strYear,strMonth,strDay
    
'    Dim iStrPrevKey

    '[[EXPORTS ���]]
    Const A505_I1_select_char = 0		'import_fg ief_supplied - ������ 
    Const A505_I2_acq_seq = 0			'import_next a_asset_acq_item - �ڻ������� 
    Const A505_I3_asst_no = 0			'import_next a_asset_master - �ڻ��ȣ 
    Const A505_I4_acq_no = 0			'import a_asset_acq - �ڻ�����ȣ 
    
    Const A311_E1_minor_nm = 0			'export b_minor - �ΰ��������� 
    Const A311_E2_gl_dt = 0				'export a_batch - ��ǥ���� 
    Const A311_E3_asst_no = 0			'export_next a_asset_master - �ڻ������� 
	Const A311_E4_acq_seq = 0			'export_next a_asset_acq_item - �ڻ��ȣ 

    'export b_acct_dept
    Const A311_E6_org_change_id = 0		'��������ID
    Const A311_E6_dept_cd = 1			'�μ��ڵ� 
    Const A311_E6_dept_nm = 2			'�μ���� 

    'export b_biz_partner
    Const A311_E7_bp_cd = 0				'�ŷ�ó�ڵ� 
    Const A311_E7_bp_type = 1			'�ŷ�óType
    Const A311_E7_bp_nm = 2				'�ŷ�ó��(���)

    'export a_asset_acq >>> �ڻ���泻�� 
    Const A311_E5_acq_no = 0			'�ڻ�����ȣ 
    Const A311_E5_acq_dt = 1			'�ڻ�������� 
    Const A311_E5_doc_cur = 2			'�ŷ���ȭ 
    Const A311_E5_xch_rate = 3			'ȯ�� 
    Const A311_E5_acq_fg = 4			'��汸�� 
    Const A311_E5_tot_acq_amt = 5		'�����ݾ� 
    Const A311_E5_tot_acq_loc_amt = 6	'�����ݾ�(�ڱ�)
    Const A311_E5_extra_acq_amt = 7		'�δ��� 
    Const A311_E5_extra_acq_loc_amt = 8	'�δ���(�ڱ�)
    Const A311_E5_vat_make_fg = 9	
    Const A311_E5_vat_no = 10			'�ΰ�����ȣ 
    Const A311_E5_vat_amt = 11			'�ΰ����ݾ� 
    Const A311_E5_vat_loc_amt = 12		'�ΰ����ݾ�(�ڱ�)
    Const A311_E5_ref_no = 13			'������ȣ 
    Const A311_E5_ap_acct_cd = 14		'�����ޱ� ���� 
    Const A311_E5_ap_no = 15			'ä����ȣ 
    Const A311_E5_ap_due_dt = 16		'ä���������� 
    Const A311_E5_ap_amt = 17			'ä���ݾ� 
    Const A311_E5_ap_loc_amt = 18		'ä���ݾ�(�ڱ�)
    Const A311_E5_gl_no = 19			'��ǥ��ȣ 
    Const A311_E5_temp_gl_no = 20		'������ǥ��ȣ 
    Const A311_E5_acq_desc = 21			'���� 
    Const A311_E5_internal_cd = 22		'���κμ��ڵ� 
    Const A311_E5_vat_type = 23			'�ΰ������� 
    Const A311_E5_vat_rate = 24			'�ΰ����� 
    
    '[[EXPORTS Group ���]]
    'Group Name : export_group		(old)
'    Const A505_EG1_E1_dept_cd = 0		'�μ��ڵ�	'b_acct_dept
'    Const A505_EG1_E1_dept_nm = 1		'�μ���� 
'    Const A505_EG1_E2_acct_cd = 2		'�����ڵ�	'a_acct
'    Const A505_EG1_E2_acct_nm = 3		'�����ܸ� 
'    Const A505_EG1_E3_asst_no = 4		'�ڻ��ȣ	'a_asset_master
 '   Const A505_EG1_E3_asst_nm = 5		'�ڻ�� 
 '   Const A505_EG1_E3_reg_dt = 7		'�ڻ������� 
 '   Const A505_EG1_E3_spec = 8			'�뵵/�԰� 
 '   Const A505_EG1_E3_doc_cur = 9		'�ŷ���ȭ 
 '   Const A505_EG1_E3_xch_rate = 10		'ȯ�� 
 '   Const A505_EG1_E3_acq_amt = 11		'���ݾ� 
 '   Const A505_EG1_E3_ref_no = 5		'������ȣ 
 '   Const A505_EG1_E3_acq_loc_amt = 12	'���ݾ�(�ڱ�)
 '   Const A505_EG1_E3_acq_qty = 13		'������ 
 '   Const A505_EG1_E3_inv_qty = 14		'������ 
 '   Const A505_EG1_E3_tax_dur_yrs = 15	'�������س����� 
 '   Const A505_EG1_E3_cas_dur_yrs = 16	'���ȸ����س����� 
 '   Const A505_EG1_E3_asset_desc = 17	'���� 
 '   Const A505_EG1_E3_gl_no = 18		'��ǥ��ȣ 
 '   Const A505_EG1_E3_temp_gl_no = 19	'������ǥ��ȣ 
 '   Const A505_EG1_E3_cas_end_l_term_depr_tot_amt = 20	'���ȸ��������⸻�����󰢴���ݾ� 
 '   Const A505_EG1_E3_start_depr_yymm = 21				'�����󰢽��۳�� 
 '   Const A505_EG1_E3_cas_depr_sts = 22	'���ȸ����ػ󰢻��� 
 
'    Const A505_EG1_E1_dept_cd = 0		'�μ��ڵ�		(new)
'    Const A505_EG1_E1_dept_nm = 1		'�μ���� 
'    Const A505_EG1_E2_acct_cd = 2		'�����ڵ� 
'    Const A505_EG1_E2_acct_nm = 3		'�����ܸ� 
'    Const A505_EG1_E3_asst_no = 4		'�ڻ��ȣ 
'    Const A505_EG1_E3_asst_nm = 5		'�ڻ�� 
'    Const A505_EG1_E3_acq_amt = 6		'���ݾ� 
'    Const A505_EG1_E3_acq_loc_amt = 7	'���ݾ�(�ڱ�)
'    Const A505_EG1_E3_acq_qty = 8		'������ 
'    Const A505_EG1_E3_res_amt = 9		'��������(�ڱ�)
'    Const A505_EG1_E3_ref_no = 10		'������ȣ 
'    Const A505_EG1_E3_asset_desc = 11	'���� 
'    Const A505_EG1_E3_reg_dt = 12		'�ڻ������� 
'    Const A505_EG1_E3_spec = 13			'�뵵/�԰� 
'    Const A505_EG1_E3_doc_cur = 14		'�ŷ���ȭ 
'    Const A505_EG1_E3_xch_rate = 15		'ȯ�� 
'    Const A505_EG1_E3_inv_qty = 16		'������ 
'    Const A505_EG1_E3_tax_dur_yrs = 17	'�������س����� 
'    Const A505_EG1_E3_cas_dur_yrs = 18	'���ȸ����س����� 
'    Const A505_EG1_E3_gl_no = 19		'��ǥ��ȣ 
'    Const A505_EG1_E3_temp_gl_no = 20	'������ǥ��ȣ 
'    Const A505_EG1_E3_cas_end_l_term_depr_tot_amt = 21		'�󰢴��� 
'    Const A505_EG1_E3_start_depr_yymm = 22					'�����󰢽��۳�� 
'    Const A505_EG1_E3_cas_depr_sts = 23					'���� 
'    Const A505_EG1_E3_cas_dur_mnth = 24	

     
    'Group Name : EG1_exp_group >>> ��ݳ��� 
'	Const A505_EG2_E1_bank_acct_no = 0    '�������ڵ� 
'	Const A505_EG2_E2_acq_seq = 1         '���� 
'	Const A505_EG2_E2_paym_type = 2		  '������� 
'	Const A505_EG2_E2_paym_amt = 3		  '�ݾ� 
'	Const A505_EG2_E2_paym_loc_amt = 4	  '�ݾ�(�ڱ�)
'	Const A505_EG2_E2_note_no = 5		  '������ȣ 
	
	
	Const C_AcqDt		= 1
    Const C_Deptcd		= 2
	Const C_DeptPop		= 3
	Const C_DeptNm		= 4
	Const C_AcctCd		= 5
	Const C_AcctPop		= 6
	Const C_AcctNm		= 7
	Const C_AsstNo		= 8
	Const C_AsstNm		= 9
    Const C_AcqAmt		= 10
	Const C_AcqLocAmt	= 11
	Const C_DeprLocAmt	= 12
	Const C_InvQty		= 13
	Const C_ResAmt		= 14
	Const C_DeprFrDt	= 15
	Const C_DurYrs		= 16
	Const C_DeprstsCd	= 17
	Const C_DeprstsPop	= 18
	Const C_Deprsts		= 19
	Const C_RefNo		= 20
	Const C_Desc		= 21



'-------------------------   
' 2.2. ��û ���� ó�� 
'-------------------------
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	Dim intMaxRows_i
	
	plgStrPrevKey_i = Request("lgStrPrevKey_i")    
	intMaxRows_i = Request("txtMaxRows_i") 
	
	'-----------------------
	'Data manipulate  area(import view match) (�ڻ�������)
	'-----------------------
	If plgStrPrevKey_i = "" Then
		I2_a_asset_acq_item = 0
	Else
		I2_a_asset_acq_item = plgStrPrevKey_i
	End If
	
	Redim I4_a_asset_acq(0)
    
    I4_a_asset_acq(A505_I4_acq_no) = Request("txtAcqNo")	'�ڻ�����ȣ 
    
	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqLkUpSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If    

	Call iPAAG010.AS0029_ACQ_LOOKUP_SVR(gStrGloBalCollection, _
										I1_ief_supplied, _
										I2_a_asset_acq_item, _
										I3_a_asset_master, _
										I4_a_asset_acq, _
										E1_b_minor, _
										E2_a_batch, _
										E3_a_asset_acq_item, _
										E4_a_asset_master, _
										EG1_exp_group, _
										E5_b_biz_partner, _
										E6_b_acct_dept, _
										E7_a_asset_acq, _
										EG2_export_itm_grp)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing

'-------------------------
' 2.4. HTML �� ��� ������ 
'-------------------------

    '�ڻ��������, ��ǥ���� 
    idate = Trim(replace(E7_a_asset_acq(A311_E5_acq_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_acq_dt) = idate

    idate = Trim(replace(E2_a_batch(A311_E2_gl_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E2_a_batch(A311_E2_gl_dt) = idate

    iStrData = ""
    For iLngRow = 0 To UBound(EG1_exp_group,1) - 1
		Call ExtractDateFrom(EG1_exp_group(iLngRow,24),"YYYYMM","",strYear,strMonth,strDay)
		StrDate = UniConvYYYYMMDDtoDate(gAPDateFormat,strYear,strMonth,"01")

		iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow,14))		'1  ������� 
		iStrData = iStrData & Chr(11) & Trim(ConvSPChars(EG1_exp_group(iLngRow,0)))			'2  �μ��ڵ� 
		iStrData = iStrData & Chr(11) & ""													'3  �μ�pop
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,1))				'4	�μ���� 
		'iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,2))				'	�ڽ�Ʈ����
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,3))				'5  �����ڵ� 
		iStrData = iStrData & Chr(11) & ""													'6  ����pop
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,4))				'7	�����ܸ� 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,5))				'8	�ڻ��ȣ 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,6))				'9	�ڻ�� 
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,7),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	'10	���ݾ� 	
		iStrData = iStrData & Chr(11) &	UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,8),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")'11	���ݾ�(�ڱ�)
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,23))				'12	�󰢴��� 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,18))				'13	������ 
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,11),lgCurrency,ggAmtOfMoneyNo, "X" , "X")'14	�������� 
		iStrData = iStrData & Chr(11) & UNIMonthClientFormat(ConvSPChars(strDate))			'15 �����󰢽��۳�� 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,26))				'16 ������� ??
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,25))				'17 �󰢻����ڵ� 
		iStrData = iStrData & Chr(11) & ""													'18 �󰢻����ڵ��˾� 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,27))				'19	�󰢻��¸� 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,12))				'20 ������ȣ 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,13))				'21 ���� 
		iStrData = iStrData & Chr(11) & intMaxRows_i + iLngRow + 1
		iStrData = iStrData & Chr(11) & Chr(12)
    Next

	plgStrPrevKey_i = ""

	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

	Response.Write "    if """ & E7_a_asset_acq(A311_E5_acq_fg) & """ <> ""03"" then " & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117214"",""X"",""X"",""X"") " & vbCr ''��汸�� üũ.
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "	else	" & vbCr

    '*** Master ***
    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
    
    '*** Control ***
    Response.Write "	.frm1.txtDeptCd.value    = """ & Trim(ConvSPChars(E6_b_acct_dept(A311_E6_dept_cd)))  & """" & vbCr '�μ��ڵ� 
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_nm))		 & """" & vbCr '�μ���� 
    Response.Write "	.frm1.txtGLDt.text       = """ & UNIDateClientFormat(E2_a_batch(A311_E2_gl_dt))		 & """" & vbCr '��ǥ���� 
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_cd))		 & """" & vbCr '�ŷ�ó�ڵ� 
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_nm))		 & """" & vbCr '�ŷ�ó��(���)
    
    Response.Write "	.frm1.txtAcqNo.value     = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_no))		 & """" & vbCr '�ڻ�����ȣ 
    Response.Write "	.frm1.txtAcqDt.text      = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_acq_dt)) & """" & vbCr '�ڻ�������� 
    Response.Write "	.frm1.txtDocCur.value    = """ & E7_a_asset_acq(A311_E5_doc_cur)					 & """" & vbCr '�ŷ���ȭ 
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_xch_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr 'ȯ�� 
    Response.Write "	.frm1.cboAcqFg.value     = """ & E7_a_asset_acq(A311_E5_acq_fg)						 & """" & vbCr '��汸�� 
    Response.Write "	.frm1.txtAcqAmt.text    = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")				 & """" & vbCr '�����ݾ� 
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr '�����ݾ�(�ڱ�)
    Response.Write "	.frm1.txtGLNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_gl_no))			 & """" & vbCr '��ǥ��ȣ 
    Response.Write "	.frm1.txtTempGLNo.value  = """ & ConvSPChars(E7_a_asset_acq(A311_E5_temp_gl_no))	 & """" & vbCr '������ǥ��ȣ 
    Response.Write "	.frm1.txtDesc.value      = """ & Trim(ConvSPChars(E7_a_asset_acq(A311_E5_acq_desc))) & """" & vbCr '���� 

    Response.Write "	.lgStrPrevKey = """ & plgStrPrevKey_i & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr

	Response.Write "    end if	" & vbCr

    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

%>
