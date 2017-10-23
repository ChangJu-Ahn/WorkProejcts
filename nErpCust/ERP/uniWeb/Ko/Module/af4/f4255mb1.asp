<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : REPAY LOAN MULTI QUERY
'*  3. Program ID        : f4255mb1
'*  4. Program �̸�      : ���Աݸ�Ƽ��ȯ(��ȸ)
'*  5. Program ����      : ���Աݸ�Ƽ��ȯ 
'*  6. Complus ����Ʈ    : PAFG430.DLL
'*  7. ���� �ۼ������   : 2003/05/10
'*  8. ���� ���������   : 2003/05/10
'*  9. ���� �ۼ���       : ����� 
'* 10. ���� �ۼ���       : ����� 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2.- ���Ǻ� 
'##########################################################################################################
																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
On Error Resume Next														'��: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################
If strMode = "" Then'
	Response.End
	Call HideStatusWnd		 
ElseIf strMode <> CStr(UID_M0001) Then										'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Call DisplayMsgBox("700118", vbOKOnly, "", "", I_MKSCRIPT)				'��ȸ��û�� �� �� �ֽ��ϴ�.
	Response.End
	Call HideStatusWnd		 
ElseIf Trim(Request("txtRePayNo")) = "" Then									'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)				'��ȸ ���ǰ��� ����ֽ��ϴ�!
	Response.End
	Call HideStatusWnd		 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim iPAFG430																'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim IntRows
Dim intCount
Dim intCount0
Dim IntCount1
Dim LngMaxRow
Dim lgCurrency
Dim txthOrgChangeId

Dim I1_f_ln_repay_pay_no

Dim E1_f_ln_repay_info
Const A860_E1_pay_no = 0
Const A860_E1_pay_dt = 1
Const A860_E1_dept_cd = 2
Const A860_E1_dept_nm = 3
Const A860_E1_pr_rdp_amt = 4
Const A860_E1_pr_rdp_loc_amt = 5
Const A860_E1_int_pay_amt = 6
Const A860_E1_int_pay_loc_amt = 7
Const A860_E1_etc_pay_amt = 8
Const A860_E1_etc_pay_loc_amt = 9
Const A860_E1_pay_amt_tot = 10
Const A860_E1_pay_loc_amt_tot = 11
Const A860_E1_user_fld1 = 12
Const A860_E1_user_fld2 = 13
Const A860_E1_repay_desc = 14
Const A860_E1_org_change_id = 15
Const A860_E1_temp_gl_no = 16
Const A860_E1_gl_no = 17

Dim EG1_f_ln_repay_mean
Const A860_EG1_mean_seq_no = 0
Const A860_EG1_mean_type = 1
Const A860_EG1_mean_type_nm = 2
Const A860_EG1_bank_acct_no = 3
Const A860_EG1_bank_cd = 4
Const A860_EG1_bank_nm = 5
Const A860_EG1_pay_mean_acct_cd = 6
Const A860_EG1_acct_nm = 7
Const A860_EG1_doc_cur = 8
Const A860_EG1_xch_rate = 9
Const A860_EG1_pay_amt = 10
Const A860_EG1_pay_loc_amt = 11
Const A860_EG1_mean_desc = 12

Dim EG2_f_ln_repay_item
Const A860_EG2_loan_no = 0
Const A860_EG2_loan_dt = 1
Const A860_EG2_due_dt = 2
Const A860_EG2_pay_plan_dt = 3
Const A860_EG2_doc_cur = 4
Const A860_EG2_xch_rate = 5
Const A860_EG2_pay_amt = 6
Const A860_EG2_pay_loc_amt = 7
Const A860_EG2_pay_dfr_amt = 8
Const A860_EG2_pay_dfr_loc_amt = 9
Const A860_EG2_pay_xch_rate = 10
Const A860_EG2_pay_int_amt = 11
Const A860_EG2_pay_int_loc_amt = 12
Const A860_EG2_plan_acct_cd = 13
Const A860_EG2_acct_nm = 14
Const A860_EG2_loan_bal_amt = 15
Const A860_EG2_loan_bal_loc_amt = 16
Const A860_EG2_rdp_amt = 17
Const A860_EG2_rdp_loc_amt = 18
Const A860_EG2_int_pay_amt = 19
Const A860_EG2_int_pay_loc_amt = 20
Const A860_EG2_item_desc = 21
Const A860_EG2_pay_obj = 22
    
Dim EG3_f_ln_repay_etc
Const A860_EG3_pay_item_acct_cd = 0
Const A860_EG3_acct_nm = 1
Const A860_EG3_pay_amt = 2
Const A860_EG3_pay_loc_amt = 3
Const A860_EG3_item_desc = 4

	I1_f_ln_repay_pay_no = Trim(Request("txtRePayNO"))

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set iPAFG430 = Server.CreateObject("PAFG430.cFLkUpRepayMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If

	Call iPAFG430.F_LOOKUP_REPAY_MULTI_SVR(gStrGlobalCollection, I1_f_ln_repay_pay_no, E1_f_ln_repay_info,  _
	                                      EG1_f_ln_repay_mean, EG2_f_ln_repay_item, EG3_f_ln_repay_etc)

	'---------------------------------------------
	'Com action result check area(OS,internal)
	'---------------------------------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG430 = Nothing																	'��: ComProxy Unload
		%><Script Language=vbscript>Parent.frm1.txtRePayNo.focus</Script><%
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If
		
	Set iPAFG430 = Nothing

	'//////////////////////////////////////////////////////////////////
	'  ���Աݸ�Ƽ��ȯ ��� ���� 
	'//////////////////////////////////////////////////////////////////
	txthOrgChangeId = ConvSPChars(E1_f_ln_repay_info(A860_E1_org_change_id))

	Response.Write "<Script Language=vbscript>"																	   & vbCr
	Response.Write " With parent.frm1 "																			   & vbCr

	Response.Write ".txtRePayNO.value		= """ & ConvSPChars(I1_f_ln_repay_pay_no)						& """" & vbCr
	Response.Write ".txtRePayDT.text		= """ & UNIDateClientFormat(E1_f_ln_repay_info(A860_E1_pay_dt))	& """" & vbCr
	Response.Write ".txtDeptCd.Value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_dept_cd))		& """" & vbCr
	Response.Write ".txtDeptNm.Value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_dept_nm))									& """" & vbCr
	Response.Write ".txtRePayTotLocAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E1_f_ln_repay_info(A860_E1_pr_rdp_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write ".txtRePayIntLocAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E1_f_ln_repay_info(A860_E1_int_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr	
	Response.Write ".txtEtcLocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_f_ln_repay_info(A860_E1_etc_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write ".txtPaymTotLocAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E1_f_ln_repay_info(A860_E1_pay_loc_amt_tot),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".txtUserFld1.value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_user_fld1))		& """" & vbCr
	Response.Write ".txtUserFld2.value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_user_fld2 ))		& """" & vbCr
	Response.Write ".txtRePayDesc.value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_repay_desc ))	& """" & vbCr							
	Response.Write ".txthTempGLNo.value		= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_temp_gl_no))		& """" & vbCr
	Response.Write ".txthGLNo.value			= """ & ConvSPChars(E1_f_ln_repay_info(A860_E1_gl_no ))			& """" & vbCr		
	
	Response.Write " End With "																					   & vbCr
    Response.Write "</Script>"																					   & vbCr

    intCount  = UBound(EG1_f_ln_repay_mean,1)
    intCount0 = UBound(EG2_f_ln_repay_item,1)
    IntCount1 = UBound(EG3_f_ln_repay_etc,1)
    
    If IntCount = "" Or IntCount = null Then
		IntCount = -1    
	End If
    
    If IntCount0 = "" Or IntCount0 = null Then
		IntCount0 = -1    
	End If
	
    If IntCount1 = "" Or IntCount1 = null Then
		IntCount1 = -1    
	End If	    
    
	'////////////////////////////////////
	'		��ݳ��� ���� 
	'////////////////////////////////////
	strData = ""

	If Not Missing(EG1_f_ln_repay_mean) And Not IsEmpty(EG1_f_ln_repay_mean) Then
		For IntRows = 0 To intCount
			lgCurrency = EG1_f_ln_repay_mean(IntRows,A860_EG1_doc_cur)

			strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_mean_seq_no))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_mean_type))
   		    strData = strData & Chr(11) & ""
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_mean_type_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_bank_acct_no))
   		    strData = strData & Chr(11) & ""   	    
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_bank_cd))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_bank_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_pay_mean_acct_cd)) 
			strData = strData & Chr(11) & ""   	    		       		      	    
			strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_acct_nm)) 
			strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_doc_cur)) 
			strData = strData & Chr(11) & ""   	    		       		      	    
		    strData = strData & Chr(11) & UNINumClientFormat(EG1_f_ln_repay_mean(IntRows,A860_EG1_xch_rate),	ggExchRate.DecPoint	,0)			
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_f_ln_repay_mean(IntRows,A860_EG1_pay_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_f_ln_repay_mean(IntRows,A860_EG1_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    strData = strData & Chr(11) & ConvSPChars(EG1_f_ln_repay_mean(IntRows,A860_EG1_mean_desc))
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
		    strData = strData & Chr(11) & Chr(12)   
		Next

		Response.Write "<Script Language=VBScript> "															& vbCr  
		Response.Write " With parent "																			& vbCr 
		Response.Write " .ggoSpread.Source = .frm1.vspdData4 "													& vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData4," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DOCCUR,.C_REPAY_AMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    .frm1.vspdData4.Redraw = True   "                      & vbCr
		Response.Write " End With "																				& vbCr
		Response.Write "</Script>"  																			& vbCr	
	End If		
	
	'////////////////////////////////////
	'		���Աݻ�ȯ ���� 
	'////////////////////////////////////
	strData = "" 

	If Not Missing(EG2_f_ln_repay_item) And Not IsEmpty(EG2_f_ln_repay_item) Then	
		For IntRows = 0 To intCount0
			lgCurrency = EG2_f_ln_repay_item(intRows,A860_EG2_doc_cur)
	
   		    strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_loan_no))
   		    strData = strData & Chr(11) & UNIDateClientFormat(EG2_f_ln_repay_item(intRows,A860_EG2_loan_dt))
   		    strData = strData & Chr(11) & UNIDateClientFormat(EG2_f_ln_repay_item(intRows,A860_EG2_due_dt))
   		    strData = strData & Chr(11) & UNIDateClientFormat(EG2_f_ln_repay_item(intRows,A860_EG2_pay_plan_dt))
   		    strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_doc_cur))
   		    strData = strData & Chr(11) & UNINumClientFormat(EG2_f_ln_repay_item(intRows,A860_EG2_xch_rate),	ggExchRate.DecPoint	,0)
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_dfr_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_dfr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNINumClientFormat(EG2_f_ln_repay_item(intRows,A860_EG2_pay_xch_rate),	ggExchRate.DecPoint	,0)
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_int_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_pay_int_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_plan_acct_cd))
			strData = strData & Chr(11) & ""   		    
			strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_acct_nm))
  		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_loan_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")			
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_loan_bal_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_rdp_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_rdp_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_int_pay_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")   		       		    
  		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_f_ln_repay_item(intRows,A860_EG2_int_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_item_desc))
			strData = strData & Chr(11) & ConvSPChars(EG2_f_ln_repay_item(intRows,A860_EG2_pay_obj))
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                 '11
		    strData = strData & Chr(11) & Chr(12)           
		Next

		Response.Write "<Script Language=VBScript> "															& vbCr  
		Response.Write " With parent "																			& vbCr 
		Response.Write " .ggoSpread.Source = .frm1.vspdData1 "													& vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_REPAY_PLAN_AMT    , ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_REPAY_INT_DFR_AMT , ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_REPAY_PLAN_INT_AMT, ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_LOAN_BAL_AMT      , ""A"" ,""I"",""X"",""X"")" & vbCr		
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_LOAN_RDP_TOT_AMT  , ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_LOAN_DOCCUR,.C_LOAN_INT_TOT_AMT  , ""A"" ,""I"",""X"",""X"")" & vbCr		
		Response.Write  "    .frm1.vspdData1.Redraw = True   "                      & vbCr
		Response.Write " End With "																				& vbCr
		Response.Write "</Script>"  																			& vbCr	
	End If    

	'////////////////////////////////////
	'		�δ��볻�� ���� 
	'////////////////////////////////////
	strData = "" 

	If Not Missing(EG3_f_ln_repay_etc) And Not IsEmpty(EG3_f_ln_repay_etc) Then	
		lgCurrency = gCurrency

		For IntRows = 0 To intCount1
   		    strData = strData & Chr(11) & ""
   		    strData = strData & Chr(11) & ConvSPChars(EG3_f_ln_repay_etc(intRows,A860_EG3_pay_item_acct_cd))
   		    strData = strData & Chr(11) & ""
   		    strData = strData & Chr(11) & ConvSPChars(EG3_f_ln_repay_etc(intRows,A860_EG3_acct_nm))
   		    strData = strData & Chr(11) & "KRW"
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_f_ln_repay_etc(intRows,A860_EG3_pay_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_f_ln_repay_etc(intRows,A860_EG3_pay_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & ConvSPChars(EG3_f_ln_repay_etc(intRows,A860_EG3_item_desc))		
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
		    strData = strData & Chr(11) & Chr(12)   
		Next

		Response.Write "<Script Language=VBScript> "					         & vbCr
		Response.Write " With parent "											 & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData                " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur,.C_ItemAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "   .frm1.vspdData.Redraw = True   "                     & vbCr
		Response.Write " End With "									 	         & vbCr
		Response.Write "</Script>"  										     & vbCr
	End If    

	Response.Write "<Script Language=VBScript> "							     & vbCr
	Response.Write " With parent "											     & vbCr
	Response.Write " .frm1.txtRePayNo.value = """ & I1_f_ln_repay_pay_no & """" & vbCr
	Response.Write " .frm1.horgChangeId.value = """ & txthOrgChangeId	 & """" & vbCr
	Response.Write " .DbQueryOk	"										         & vbCr
    Response.Write " End With "									 		         & vbCr
    Response.Write "</Script>"  										         & vbCr

%>	
	
