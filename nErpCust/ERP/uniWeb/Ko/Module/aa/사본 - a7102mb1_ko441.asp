<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

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
'					1. Include
'===================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'===================================================================
'					2. ���Ǻ� 
'===================================================================
	On Error Resume Next
	Err.Clear  

	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide


    '---------------------------------------Common-----------------------------------------------------------
    Call LoadBasisGlobalInf()
	Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
'    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'===================================================================
'					2.1 ���� üũ 
'===================================================================

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
'					2. ���� ó�� ����� 
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
    Dim E8_b_tax_biz_area	'���ݽŰ�����
	'20030301	�����ޱݰ����߰�
	Dim E9_a_acct			'�����ޱݰ���
	'20050512	�ſ�ī���ȣ �߰�
	DIM E10_CREDIT_CARD_NO	'�ſ�ī�� ��ȣ
  
    Dim iLngRow,iLngCol
    Dim idate
    Dim iStrData
    Dim iStrData2
    Dim intMaxRows_i
    Dim intMaxRows_m
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
    Const A311_E5_ref_no = 13			'ǰ��׷�
    '20030301 �����ޱݰ��� �߰� -> A311_E9_ap_acct_cd �� ��ü.
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
    
'�߰�����
    Const A311_E5_issued_dt = 25
    Const A311_E5_tax_biz_area_cd = 26

	'[CONVERSION INFORMATION]  EXPORTS View ���
	'[CONVERSION INFORMATION]  View Name : exp_fr b_tax_biz_area
	Const A311_E8_tax_biz_area_cd = 0
	Const A311_E8_tax_biz_area_nm = 1
'�߰�����
 
    '20030301   EXPORTS View ���
    Const A311_E9_ap_acct_cd = 0
    Const A311_E9_ap_acct_nm = 1
    
    '20050512	EXPORTS View ���
    Const A311_E10_CREDIT_CARD_NO = 0
    Const A311_E10_CREDIT_CARD_NM = 1
    

    '[[EXPORTS Group ���]]
    'Group Name : export_group(b_acct_dept, a_acct,a_asset_master) >>> �ڻ�Master
    Const A311_EG2_E1_dept_cd = 0		'�μ��ڵ�
    Const A311_EG2_E1_dept_nm = 1		'�μ����
    Const A311_EG2_E3_cost_cd = 3		'Cost Center : air   
    Const A311_EG2_E2_acct_cd = 4		'�����ڵ�
    Const A311_EG2_E2_acct_nm = 5		'�����ܸ�
    Const A311_EG2_E3_asst_no = 6		'�ڻ��ȣ
    Const A311_EG2_E3_asst_nm = 7		'�ڻ��
    Const A311_EG2_E3_acq_amt = 8		'���ݾ�
    Const A311_EG2_E3_acq_loc_amt = 9	'���ݾ�(�ڱ�)
    Const A311_EG2_E3_acq_qty = 10		'������
    Const A311_EG2_E3_dur_yrs = 11		'���뿬��    : air 
    Const A311_EG2_E3_res_amt = 12		'��������(�ڱ�)
    Const A311_EG2_E3_ref_no = 12		'ǰ��׷�	 : air
    Const A311_EG2_E3_asset_desc = 13	'����

    Const A311_EG2_E3_reg_dt = 14		'�ڻ�������
    Const A311_EG2_E3_spec = 15			'�뵵/�԰�
    Const A311_EG2_E3_doc_cur = 16		'�ŷ���ȭ
    Const A311_EG2_E3_xch_rate = 17		'ȯ��
    Const A311_EG2_E3_inv_qty = 18		'������
    Const A311_EG2_E3_tax_dur_yrs = 19	'�������س�����
    Const A311_EG2_E3_cas_dur_yrs = 20	'���ȸ����س�����
    Const A311_EG2_E3_gl_no = 21		'��ǥ��ȣ
    Const A311_EG2_E3_temp_gl_no = 22	'������ǥ��ȣ
    
    'Group Name : EG1_exp_group >>> ��ݳ���
	Const A505_EG2_E1_bank_acct_no = 0    '�������ڵ�
	Const A505_EG2_E2_acq_seq = 1         '����
	Const A505_EG2_E2_paym_type = 2		  '�������
	Const A505_EG2_E2_paym_amt = 3		  '�ݾ�
	Const A505_EG2_E2_paym_loc_amt = 4	  '�ݾ�(�ڱ�)
	Const A505_EG2_E2_note_no = 5		  '������ȣ
	Const A505_EG2_E2_b_minor_nm = 6	  '�����ݸ�
	Const A505_EG2_E2_bulid_asst_no = 7	  '�Ǽ������ڻ��ȣ '>>air


	' -- ���Ѱ����߰�
	Const I5_a_data_auth_data_BizAreaCd = 0
	Const I5_a_data_auth_data_internal_cd = 1
	Const I5_a_data_auth_data_sub_internal_cd = 2
	Const I5_a_data_auth_data_auth_usr_id = 3

	Dim I5_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ����
	
'-------------------------   
' 2.2. ��û ���� ó�� 
'-------------------------
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	plgStrPrevKey_i = Request("lgStrPrevKey_i")     
	intMaxRows_i = Request("txtMaxRows_i")
	intMaxRows_m = Request("txtMaxRows_m")	

  	Redim I5_a_data_auth(3)
	I5_a_data_auth(I5_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I5_a_data_auth(I5_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I5_a_data_auth(I5_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I5_a_data_auth(I5_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
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
    
	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqLkUpSvr")	'AIR :: PAAG010 -> PAAG010_KO441

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
										EG2_export_itm_grp, _
										E8_b_tax_biz_area, _
										E9_a_acct, _
										E10_CREDIT_CARD_NO, _
										I5_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then			
		
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing

'-------------------------
' 2.4. HTML �� ��� ������ 
'-------------------------

    '�ڻ��������, �����ޱ� ��������, ��ǥ����
    idate = Trim(replace(E7_a_asset_acq(A311_E5_ap_due_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_ap_due_dt) = idate

    idate = Trim(replace(E7_a_asset_acq(A311_E5_acq_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_acq_dt) = idate

    idate = Trim(replace(E2_a_batch(A311_E2_gl_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E2_a_batch(A311_E2_gl_dt) = idate

	iStrData = ""
    For iLngRow = 0 To UBound(EG1_exp_group,1) - 1
        For iLngCol = 0 To UBound(EG1_exp_group,2)
        	if iLngCol = A311_EG2_E1_dept_cd then			'�μ��ڵ� 0		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
        	elseif iLngCol = A311_EG2_E1_dept_nm then       '�μ���� 1		
                	iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))

 	
        	elseif iLngCol = A311_EG2_E3_cost_cd then       '�ڽ�Ʈ����	2	: air
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
        		
        	elseif iLngCol = A311_EG2_E2_acct_cd then		'�����ڵ� 3		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep


        	elseif iLngCol = A311_EG2_E2_acct_nm then       '�����ܸ� 4		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_asst_no then       '�ڻ��ȣ 5		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_asst_nm then       '�ڻ��   6		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_acq_amt then       '���ݾ� 7		
        		iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		elseif iLngCol = A311_EG2_E3_acq_loc_amt then	'���ݾ�(�ڱ�) 8  
				iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		elseif iLngCol = A311_EG2_E3_acq_qty then       '������ 9       
			iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
		
		elseif iLngCol = A311_EG2_E3_dur_yrs then		'���뿬�� 10	: air 	
			iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		
		elseif iLngCol = A311_EG2_E3_res_amt then       '��������(�ڱ�) 11 
			iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),lgCurrency,ggAmtOfMoneyNo, "X" , "X")			
		elseif iLngCol = A311_EG2_E3_ref_no then        'ǰ��׷� 12	: air
		       iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
		elseif iLngCol = A311_EG2_E3_asset_desc then    '���� 13		
		     iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))



        	end if
        	'iLngCol & "-" & 
		next
		
		iStrData = iStrData & gColSep & intMaxRows_i + iLngRow + 1
		iStrData = iStrData & gColSep & gRowSep
    Next
	'--------------------------------------------------
	'Spread Column
	'Const C_Seq		   = 1			'����
	'Const C_RcptType	   = 2			'�������
	'Const C_RcptTypeNm	   = 3			'���������
	'Const C_Amt		   = 4			'�ݾ�
	'Const C_LocAmt		   = 5			'�ݾ�(�ڱ�)
	'Const C_BankAcct	   = 6			'�������ڵ�
	'Const C_BankAcctPopup = 7
	'Const C_NoteNo		   = 8			'������ȣ
	'Const C_NoteNoPopup   = 9
'    Public Const A505_EG2_E1_bank_acct_no = 0    'View Name : exp_itm_item b_bank_acct
'    Public Const A505_EG2_E2_acq_seq = 1         'View Name : exp_itm_item a_asset_acq_item
'    Public Const A505_EG2_E2_paym_type = 2
'    Public Const A505_EG2_E2_paym_amt = 3
'    Public Const A505_EG2_E2_paym_loc_amt = 4
'    Public Const A505_EG2_E2_note_no = 5
	'--------------------------------------------------

	if isarray(EG2_export_itm_grp) then
		iStrData2 = ""
		For iLngRow = 0 To UBound(EG2_export_itm_grp,1) - 1

			iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_acq_seq))
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_type))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_b_minor_nm))
		    iStrData2 = iStrData2 & gColSep & UNIConvNumDBToCompanyByCurrency(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData2 = iStrData2 & gColSep & UNIConvNumDBToCompanyByCurrency(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_loc_amt),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E1_bank_acct_no))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_note_no))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_bulid_asst_no))
		    iStrData2 = iStrData2 & gColSep & ""		    
			iStrData2 = iStrData2 & gColSep & intMaxRows_m + iLngRow + 1
			iStrData2 = iStrData2 & gColSep & gRowSep
		Next
	end if

	plgStrPrevKey_i = ""

	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

	Response.Write "    if """ & E7_a_asset_acq(A311_E5_acq_fg) & """ = ""03"" then " & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117217"",""X"",""X"",""X"") " & vbCr ''��汸�� üũ.
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "	else	" & vbCr
    
    '*** Master ***
    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
    
    '*** ��ݳ��� ***
	Response.Write "	.ggoSpread.Source = .frm1.vspdData2 " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData2 & """" & vbCr

    '*** Control ***
    Response.Write "	.frm1.txtDeptCd.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_cd))			 & """" & vbCr '�μ��ڵ�            
    Response.Write "	.frm1.horgchangeID.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_org_change_id))			 & """" & vbCr '�μ��ڵ�            

    
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_nm))			 & """" & vbCr '�μ����        
    Response.Write "	.frm1.txtGLDt.text       = """ & UNIDateClientFormat(E2_a_batch(A311_E2_gl_dt))			 & """" & vbCr '��ǥ����            
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_cd))			 & """" & vbCr '�ŷ�ó�ڵ� 
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_nm))			 & """" & vbCr '�ŷ�ó��(���)           
    Response.Write "	.frm1.txtVatTypeNm.value = """ & ConvSPChars(E1_b_minor(A311_E1_minor_nm))		         & """" & vbCr '�ΰ���������        
    
    Response.Write "	.frm1.txtAcqNo.value     = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_no))			 & """" & vbCr '�ڻ�����ȣ        
    Response.Write "	.frm1.txtAcqDt.text      = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_acq_dt))	 & """" & vbCr '�ڻ��������        
    Response.Write "	.frm1.txtDocCur.value    = """ & E7_a_asset_acq(A311_E5_doc_cur)						 & """" & vbCr '�ŷ���ȭ            
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_xch_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr 'ȯ��                
    Response.Write "	.frm1.cboAcqFg.value     = """ & E7_a_asset_acq(A311_E5_acq_fg)							 & """" & vbCr '��汸��            
    Response.Write "	.frm1.txtAcqAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")				 & """" & vbCr '�����ݾ�          
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr '�����ݾ�(�ڱ�)    
    Response.Write "	.frm1.txtVatAmt.value    = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_vat_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					 & """" & vbCr '�ΰ����ݾ�          
    Response.Write "	.frm1.txtVatLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	 & """" & vbCr '�ΰ����ݾ�(�ڱ�)    
    '20030301 �����ޱݰ��� �߰�
    Response.Write "	.frm1.txtApAcctCd.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_cd))				 & """" & vbCr '�����ޱݰ���            
    Response.Write "	.frm1.txtApAcctNm.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_nm))				 & """" & vbCr '�����ޱݰ��� 
    '20050512 �ſ�ī�� ��ȣ �߰�     
    Response.Write "	.frm1.txtCardNo.value  = """ & ConvSPChars(E10_CREDIT_CARD_NO(A311_E10_CREDIT_CARD_NO))				 & """" & vbCr '�ſ�ī���ȣ            
    Response.Write "	.frm1.txtCardNm.value  = """ & ConvSPChars(E10_CREDIT_CARD_NO(A311_E10_CREDIT_CARD_NM))				 & """" & vbCr '�ſ�ī��� 
          
    Response.Write "	.frm1.txtApNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_ap_no))				 & """" & vbCr 'ä����ȣ            
    Response.Write "	.frm1.txtApDueDt.text    = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_ap_due_dt))	 & """" & vbCr 'ä����������        
    Response.Write "	.frm1.txtApAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_ap_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					 & """" & vbCr 'ä���ݾ�            
    Response.Write "	.frm1.txtApLocAmt.value  = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_ap_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		 & """" & vbCr 'ä���ݾ�(�ڱ�)      
    Response.Write "	.frm1.txtGLNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_gl_no))				 & """" & vbCr '��ǥ��ȣ            
    Response.Write "	.frm1.txtTempGLNo.value  = """ & ConvSPChars(E7_a_asset_acq(A311_E5_temp_gl_no))		 & """" & vbCr '������ǥ��ȣ        
    Response.Write "	.frm1.txtDesc.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_desc))			 & """" & vbCr '����                
    Response.Write "	.frm1.txtVatType.value   = """ & Trim(ConvSPChars(E7_a_asset_acq(A311_E5_vat_type)))			 & """" & vbCr '�ΰ�������          
    Response.Write "	.frm1.txtVatRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_vat_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr '�ΰ�����            



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'10 �� ���� ��ġ �߰� ����
	Response.Write "	.frm1.txtReportAreaCd.value        = """ & ConvSPChars(E8_b_tax_biz_area(A311_E8_tax_biz_area_cd)) &				"""" & vbCr
	Response.Write "	.frm1.txtReportAreaNm.value        = """ & ConvSPChars(E8_b_tax_biz_area(A311_E8_tax_biz_area_nm)) &				"""" & vbCr    		 	    
	Response.Write "	.frm1.fpDateTime4.text				= """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_issued_dt)) &	"""" & vbCr       'AP ��������       '��������        
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    Response.Write "	.lgStrPrevKey = """ & plgStrPrevKey_i & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
	
	
	Response.Write "    end if	" & vbCr
	
	
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

%>
