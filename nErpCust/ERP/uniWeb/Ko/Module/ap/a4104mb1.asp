<%
Option Explicit		
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : A404MB1
'*  4. Program Name         : PAYMENT ��ȸ�ϴ� P/G
'*  5. Program Desc         : PAYMENT ��ȸ�ϴ� P/G
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : CHANG SUNG HEE
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. ���Ǻ� 
'##########################################################################################################
																					'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd													

On Error Resume Next																'��: 
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																			'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")														'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then												'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Call DisplayMsgBox("700118", vbOKOnly, "", "", I_MKSCRIPT)						'��: ��ȸ �����ε� �ٸ� ���·� ��û�� ���� ���, �ʿ������ ���� ��, �޼����� ID������ ����ؾ� �� 
	Response.End 
ElseIf Request("txtAllcNo") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)						'��:
	Response.End 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim IntRows
Dim intCount
Dim IntCount1
Dim iPAPG020																		'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim lgCurrency
Dim LngMaxRow
Dim LngMaxRow1
Dim strData, strData1

'#########################################################################################################
'												2.4. HTML ��� ������ 
'##########################################################################################################

Dim  I1_a_allc_paym_no 

Dim  E1_a_allc_paym 
Const A294_E1_paym_no = 0
Const A294_E1_paym_dt = 1
Const A294_E1_bp_cd = 2
Const A294_E1_bp_nm = 3
Const A294_E1_dept_cd = 4
Const A294_E1_dept_nm = 5
Const A294_E1_org_change_id = 6
Const A294_E1_paym_type = 7
Const A294_E1_paym_type_nm = 8
Const A294_E1_bank_cd = 9
Const A294_E1_bank_nm = 10
Const A294_E1_bank_acct_no = 11
Const A294_E1_note_no = 12
Const A294_E1_acct_cd = 13
Const A294_E1_acct_nm = 14
Const A294_E1_temp_gl_no = 15
Const A294_E1_gl_no = 16
Const A294_E1_doc_cur = 17
Const A294_E1_xch_rate = 18
Const A294_E1_paym_amt = 19
Const A294_E1_paym_loc_amt = 20
Const A294_E1_dc_amt = 21
Const A294_E1_dc_loc_amt = 22
Const A294_E1_paym_desc = 23

Dim  EG1_export_group 
Const A294_EG1_a_open_ap_ap_no = 0
Const A294_EG1_a_open_ap_ap_dt = 1
Const A294_EG1_a_open_ap_ap_due_dt = 2
Const A294_EG1_a_open_ap_doc_cur = 3
Const A294_EG1_a_open_ap_xch_rate = 4
Const A294_EG1_a_open_ap_ap_amt = 5
Const A294_EG1_a_open_ap_bal_amt = 6
Const A294_EG1_a_cls_ap_cls_amt = 7
Const A294_EG1_a_cls_ap_cls_loc_amt = 8
Const A294_EG1_a_cls_ap_dc_amt = 9
Const A294_EG1_a_cls_ap_dc_loc_amt = 10
Const A294_EG1_a_cls_ap_cls_ap_desc = 11
Const A294_EG1_a_open_ap_acct_cd = 12
Const A294_EG1_a_acct_acct_nm = 13
Const A294_EG1_a_open_ap_biz_area_cd = 14
Const A294_EG1_b_biz_area_biz_area_nm = 15

Dim  EG2_export_group_dc 
Const A294_EG2_E2_a_paym_dc_seq = 0
Const A294_EG2_E1_a_paym_dc_acct_cd = 1
Const A294_EG2_E1_a_acct_acct_nm = 2
Const A294_EG2_E2_a_paym_dc_dc_amt = 3
Const A294_EG2_E2_a_paym_dc_dc_loc_amt = 4

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Const A294_I1_paym_no = 0

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

    Redim I1_a_allc_paym_no(A294_I1_paym_no+4)
    I1_a_allc_paym_no(A294_I1_paym_no)   = Trim(Request("txtAllcNo"))
	I1_a_allc_paym_no(A294_I1_paym_no+1) = lgAuthBizAreaCd
	I1_a_allc_paym_no(A294_I1_paym_no+2) = lgInternalCd
	I1_a_allc_paym_no(A294_I1_paym_no+3) = lgSubInternalCd
	I1_a_allc_paym_no(A294_I1_paym_no+4) = lgAuthUsrID	

	Set iPAPG020 = Server.CreateObject("PAPG020.cALkUpPayAllcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	    
	Call iPAPG020.A_LOOKUP_ALLC_PAYM_SVR (gStrGlobalCollection, I1_a_allc_paym_no ,E1_a_allc_paym,EG1_export_group,EG2_export_group_dc)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG020 = Nothing
		Response.End 
	End If

	Set iPAPG020 = Nothing
	
    intCount = UBound(EG1_export_group,1)
	intCount1 =  UBound(EG2_export_group_dc,1)    
    
    If IntCount = "" Or IntCount = null Then
		IntCount = -1    
	End If
    
    If IntCount1 = "" Or IntCount1 = null Then
		IntCount1 = -1    
	End If	    
    	
	lgCurrency = ConvSPChars(E1_a_allc_paym(A294_E1_doc_cur))
	LngMaxRow =  CLng(Request("txtMaxRows"))
	LngMaxRow1 =  CLng(Request("txtMaxRows1"))
	
    Response.Write "<Script Language=VBScript> " & vbCr
	Response.Write " With parent.frm1 "          & vbCr
	Response.Write " .hApDocCur.value  = """ & ConvSPChars(EG1_export_group(0,A294_EG1_a_open_ap_doc_cur)) & """" & vbCr
    '-----------------------
	'Result data display area
	'-----------------------
	
	Response.Write ".txtAllcDt.TEXT				= """ & UNIDateClientFormat(E1_a_allc_paym(A294_E1_paym_dt)) & """" & vbCr
	Response.Write ".txtDeptCd.Value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_dept_cd))         & """" & vbCr
    Response.Write ".txtDeptNm.Value		    = """ & ConvSPChars(E1_a_allc_paym(A294_E1_dept_nm))         & """" & vbCr
    Response.Write ".txtBankCd.Value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_bank_cd))         & """" & vbCr
    Response.Write ".txtBankNm.Value		    = """ & ConvSPChars(E1_a_allc_paym(A294_E1_bank_nm))         & """" & vbCr
    Response.Write ".txtBpCd.Value				= """ & ConvSPChars(E1_a_allc_paym(A294_E1_bp_cd))           & """" & vbCr
    Response.Write ".txtBpNm.Value				= """ & ConvSPChars(E1_a_allc_paym(A294_E1_bp_nm))           & """" & vbCr
    Response.Write ".txtBankAcct.Value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_bank_acct_no ))   & """" & vbCr
    Response.Write ".txtInputType.Value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_paym_type))       & """" & vbCr
    Response.Write ".txtInputTypeNm.Value		= """ & ConvSPChars(E1_a_allc_paym(A294_E1_paym_type_nm))    & """" & vbCr
    Response.Write ".txtCheckCd.Value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_note_no))         & """" & vbCr
    Response.Write ".txtDocCur.value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_doc_cur))         & """" & vbCr
    
    Response.Write ".txtGlNo.value				= """ & ConvSPChars(E1_a_allc_paym(A294_E1_gl_no))           & """" & vbCr
    Response.Write ".txtTempGlNo.value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_temp_gl_no))      & """" & vbCr
    
    Response.Write ".txtXchRate.Text			= """ & UNINumClientFormat(E1_a_allc_paym(A294_E1_xch_rate), ggAmtOfMoney.DecPoint, 0)                                        & """" & vbCr
    Response.Write ".txtPaymAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_paym(A294_E1_paym_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")               & """" & vbCr
    Response.Write ".txtPaymLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_paym(A294_E1_paym_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
    Response.Write ".txtDcAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_paym(A294_E1_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
    Response.Write ".txtDcLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_paym(A294_E1_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
    Response.Write ".txtPaymDesc.value			= """ & ConvSPChars(E1_a_allc_paym(A294_E1_paym_desc))																		  & """" & vbCr
      
	Response.Write " End With "                 & vbCr
    Response.Write "</Script> "                 & vbCr	
	 
	lgCurrency = ConvSPChars(EG1_export_group(0,A294_EG1_a_open_ap_doc_cur))
	
	If Not Missing(EG1_export_group) And Not IsEmpty(EG1_export_group) Then	
		For IntRows = 0 To intCount		
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_open_ap_ap_no))
		    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A294_EG1_a_open_ap_ap_dt))     
		    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A294_EG1_a_open_ap_ap_due_dt))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_open_ap_doc_cur))           
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_open_ap_ap_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_open_ap_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_cls_ap_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_cls_ap_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_cls_ap_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A294_EG1_a_cls_ap_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")                
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_cls_ap_cls_ap_desc))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_open_ap_acct_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_acct_acct_nm))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_a_open_ap_biz_area_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A294_EG1_b_biz_area_biz_area_nm))
		    strData = strData & Chr(11) & LngMaxRow1 + IntRows         
			strData = strData & Chr(11) & Chr(12)                                    
		Next
	End If
	
	If Not Missing(EG2_export_group_dc) And Not IsEmpty(EG2_export_group_dc) Then		  
		For IntRows = 0 To intCount1
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A294_EG2_E2_a_paym_dc_seq))
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A294_EG2_E1_a_paym_dc_acct_cd))
		    strData1 = strData1 & Chr(11) & ""
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A294_EG2_E1_a_acct_acct_nm))
			    
		    strData1 = strData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(IntRows,A294_EG2_E2_a_paym_dc_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData1 = strData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(IntRows,A294_EG2_E2_a_paym_dc_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			    	    
		    strData1 = strData1 & Chr(11) & LngMaxRow + IntRows                                 
		    strData1 = strData1 & Chr(11) & Chr(12)           
		Next
	End If
		
	Response.Write "<Script Language=VBScript> "														& vbCr  
    Response.Write " With parent "																		& vbCr
    Response.Write " .ggoSpread.Source      = .frm1.vspdData1 "											& vbCr
    Response.Write " .ggoSpread.SSShowData       """ & strData									 & """" & vbCr 
    Response.Write " .ggoSpread.Source      = .frm1.vspdData "											& vbCr
    Response.Write " .ggoSpread.SSShowData        """ & strData1								 & """" & vbCr
    Response.Write " .DbQueryOk "														 				& vbCr   
	Response.Write ".frm1.txtAcctCd.Value		= """ & ConvSPChars(E1_a_allc_paym(A294_E1_acct_cd)) & """" & vbCr
	Response.Write ".frm1.txtAcctNm.Value		= """ & ConvSPChars(E1_a_allc_paym(A294_E1_acct_nm)) & """" & vbCr    
    Response.Write " End With "                                                                         & vbCr
    Response.Write "</Script>"  																		& vbCr          

%>
