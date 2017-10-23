<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : Multi Alloction Query
'*  3. Program ID        : a3125mb1
'*  4. Program �̸�      : ��Ƽ�Ա�(��ȸ)
'*  5. Program ����      : ��Ƽ�Ա� ��ȸ 
'*  6. Complus ����Ʈ    : PARG060
'*  7. ���� �ۼ������   : 2003/03/25
'*  8. ���� ���������   : 2003/03/25
'*  9. ���� �ۼ���       : ����� 
'* 10. ���� �ۼ���       : ����� 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************

														'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
														'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


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

If Trim(Request("lgStrPrevKey")) = "" Then
	lgStrPrevKey = ""
Else
	lgStrPrevKey = Trim(Request("lgStrPrevKey"))
End If

If Trim(Request("lgStrPrevKey1")) = "" Then
	lgStrPrevKey1 = ""
Else
	lgStrPrevKey1 = Trim(Request("lgStrPrevKey1"))
End If

If Trim(Request("lgStrPrevKeyDtl")) = "" Then
	lgStrPrevKeyDtl = ""
Else
	lgStrPrevKeyDtl = Trim(Request("lgStrPrevKeyDtl"))
End If

 
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
ElseIf Trim(Request("txtRcptNo")) = "" Then									'��: ��ȸ�� ���� ���� ���Դ��� üũ 
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
Dim iPARG060																'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim IntRows
Dim IntDtlRows
Dim IntCols
Dim sList
Dim strData1
Dim strData2
Dim vbIntRet
Dim intCount
Dim IntCount1
Dim IntCurSeq
Dim LngMaxRow
Dim StrNextKey
Dim StrNextKeyDtl
Dim lgStrPrevKey
Dim lgStrPrevKeyDtl
Dim lgIntFlgMode
Dim TempInv_dt
Dim Tempbl_dt
Dim lgCurrency
Dim txthOrgChangeId

Dim I1_a_rcpt_no 

Dim E1_a_allc_rcpt 
Const A290_E1_allc_rcpt_allc_dt = 0
Const A290_E1_allc_rcpt_bp_cd = 1
Const A290_E1_allc_rcpt_bp_nm = 2
Const A290_E1_allc_rcpt_org_change_id = 3
Const A290_E1_allc_rcpt_dept_cd = 4
Const A290_E1_allc_rcpt_dept_nm = 5
Const A290_E1_allc_rcpt_temp_gl_no = 6
Const A290_E1_allc_rcpt_gl_no = 7
Const A290_E1_allc_rcpt_desc = 8

Dim E2_a_sum_amt 
Const A290_E2_allc_amt_tot = 0
Const A290_E2_cls_amt_tot = 1
Const A290_E2_etc_dr_amt_tot = 2
Const A290_E2_etc_cr_amt_tot = 3
Const A290_E2_differ_amt = 4

Dim EG1_export_group_allc     
Const A290_EG1_E1_item_seq = 0
Const A290_EG1_E2_rcpt_type = 1
Const A290_EG1_E2_rcpt_type_nm = 2
Const A290_EG1_E2_etc_no = 3
Const A290_EG1_E3_bp_cd = 4
Const A290_EG1_E3_bp_nm = 5
Const A290_EG1_E3_doc_cur = 6
Const A290_EG1_E3_xch_rate = 7
Const A290_EG1_E4_bal_amt = 8
Const A290_EG1_E4_bal_loc_amt = 9
Const A290_EG1_E4_allc_amt = 10
Const A290_EG1_E4_allc_loc_amt = 11
Const A290_EG1_E5_item_desc = 12
Const A290_EG1_E5_biz_area_cd = 13
Const A290_EG1_E5_biz_area_nm = 14
Const A290_EG1_E5_acct_cd = 15
Const A290_EG1_E5_acct_nm = 16
Const A290_EG1_E5_bank_cd = 17
Const A290_EG1_E5_bank_nm = 18

Dim EG2_export_group_cls     
Const A290_EG2_E1_ar_no = 0
Const A290_EG2_E2_ar_due_dt = 1
Const A290_EG2_E2_pay_bp_nm = 2
Const A290_EG2_E2_doc_cur = 3
Const A290_EG2_E2_xch_rate = 4
Const A290_EG2_E3_bal_amt = 5
Const A290_EG2_E3_bal_loc_amt = 6
Const A290_EG2_E3_cls_amt = 7
Const A290_EG2_E3_cls_loc_amt = 8
Const A290_EG2_E4_cls_desc = 9
Const A290_EG2_E5_ar_amt = 10
Const A290_EG2_E5_ar_loc_amt = 11
Const A290_EG2_E5_ar_dt = 12
Const A290_EG2_E5_dc_amt = 13
Const A290_EG2_E5_dc_loc_amt = 14
Const A290_EG2_E5_dc_type = 15

Dim EG3_export_group_dc     
Const A290_EG3_E1_seq = 0
Const A290_EG3_E2_acct_cd = 1
Const A290_EG3_E2_acct_nm = 2
Const A290_EG3_E2_dr_cr_fg = 3
Const A290_EG3_E2_doc_cur = 4
Const A290_EG3_E2_xch_rate = 5
Const A290_EG3_E3_dc_amt = 6
Const A290_EG3_E3_dc_loc_amt = 7
Const A290_EG3_E3_dc_desc = 8


	I1_a_rcpt_no = Trim(Request("txtRCPTNO"))

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set iPARG060 = Server.CreateObject("PARG060.cALkUpAllcMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If

	Call iPARG060.A_LOOKUP_ALLC_RCPT_MULTI_SVR(gStrGlobalCollection, I1_a_rcpt_no, E1_a_allc_rcpt,E2_a_sum_amt, _
	                                      EG1_export_group_allc, EG2_export_group_cls, EG3_export_group_dc)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG060 = Nothing																	'��: ComProxy Unload
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If
		
	Set iPARG060 = Nothing

	'//////////////////////////////////////////////////////////////////
	'  ��Ƽ�Ա� ��� ���� 
	'//////////////////////////////////////////////////////////////////
	txthOrgChangeId = ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_org_change_id))


	Response.Write "<Script Language=vbscript>"																		   & vbCr
	Response.Write " With parent.frm1 "																				   & vbCr

	Response.Write ".txtRcptNo.value	 = """ & ConvSPChars(I1_a_rcpt_no)										& """" & vbCr
	Response.Write ".txtRcptDt.text	 = """ & UNIDateClientFormat(E1_a_allc_rcpt(A290_E1_allc_rcpt_allc_dt))	& """" & vbCr
	Response.Write ".txtBpCd.Value		 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_bp_cd))			& """" & vbCr
	Response.Write ".txtBpNm.Value		 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_bp_nm))			& """" & vbCr
	Response.Write ".txtDeptCd.Value	 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_dept_cd))			& """" & vbCr
	Response.Write ".txtDeptNm.Value	 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_dept_nm))									& """" & vbCr
	Response.Write ".txtTotLocAmt.text	 = """ & UNIConvNumDBToCompanyByCurrency(E2_a_sum_amt(A290_E2_differ_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	  & """" & vbCr
	Response.Write ".txtAllcLocAmt.text	 = """ & UNIConvNumDBToCompanyByCurrency(E2_a_sum_amt(A290_E2_allc_amt_tot),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	  & """" & vbCr	
	Response.Write ".txtArClsLocAmt.text = """ & UNIConvNumDBToCompanyByCurrency(E2_a_sum_amt(A290_E2_cls_amt_tot),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	  & """" & vbCr
	Response.Write ".txtDrLocAmt.text	 = """ & UNIConvNumDBToCompanyByCurrency(E2_a_sum_amt(A290_E2_etc_dr_amt_tot),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".txtCrLocAmt.text	 = """ & UNIConvNumDBToCompanyByCurrency(E2_a_sum_amt(A290_E2_etc_cr_amt_tot),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr	
	Response.Write ".txtTempGLNo.value	 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_temp_gl_no))		& """" & vbCr
	Response.Write ".txtGLNo.value		 = """ & ConvSPChars(E1_a_allc_rcpt(A290_E1_allc_rcpt_gl_no ))			& """" & vbCr		
	
	Response.Write " End With "																						   & vbCr
    Response.Write "</Script>"																						   & vbCr

    intCount = UBound(EG1_export_group_allc,1)
    intCount0 = UBound(EG2_export_group_cls,1)
    IntCount1 = UBound(EG3_export_group_dc,1)
    
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
	'		��Ƽ�������� ���� 
	'////////////////////////////////////
	strData = ""

	If Not Missing(EG1_export_group_allc) And Not IsEmpty(EG1_export_group_allc) Then
		For IntRows = 0 To intCount
			lgCurrency = EG1_export_group_allc(IntRows,A290_EG1_E3_doc_cur)

   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E2_rcpt_type))
   		    strData = strData & Chr(11) & ""   	    
   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E2_rcpt_type_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E2_etc_no))
   		    strData = strData & Chr(11) & ""   	    
   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E3_bp_cd))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E3_bp_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E3_doc_cur)) 
			strData = strData & Chr(11) & ""   	       		      	    
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_allc(IntRows,A290_EG1_E3_xch_rate),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_allc(IntRows,A290_EG1_E4_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_allc(IntRows,A290_EG1_E4_allc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_allc(IntRows,A290_EG1_E4_allc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_acct_cd))
   		    strData = strData & Chr(11) & ""   	    		    
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_acct_nm))   		    
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_item_desc))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_bank_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_bank_nm))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_biz_area_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_allc(IntRows,A290_EG1_E5_biz_area_nm))                
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
		    strData = strData & Chr(11) & Chr(12)   
		Next

		Response.Write "<Script Language=VBScript> "															& vbCr  
		Response.Write " With parent "																			& vbCr 
		Response.Write " .ggoSpread.Source = .frm1.vspdData4 "													& vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData4," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur4,.C_BALAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData4," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur4,.C_ALLCAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    .frm1.vspdData4.Redraw = True   "                      & vbCr
		Response.Write " End With "																				& vbCr
		Response.Write "</Script>"  																			& vbCr	
	End If		
	
	'////////////////////////////////////
	'		ä�ǹ������� ���� 
	'////////////////////////////////////
	strData = "" 

	If Not Missing(EG2_export_group_cls) And Not IsEmpty(EG2_export_group_cls) Then	
		For IntRows = 0 To intCount0
			lgCurrency = EG2_export_group_cls(intRows,A290_EG2_E2_doc_cur)
	
   		    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls(intRows,A290_EG2_E1_ar_no))
   		    strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group_cls(intRows,A290_EG2_E2_ar_due_dt))
   		    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls(intRows,A290_EG2_E2_pay_bp_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls(intRows,A290_EG2_E2_doc_cur))
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E2_xch_rate),lgCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E3_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E3_cls_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E3_cls_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E5_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E5_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls(intRows,A290_EG2_E5_dc_type))
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls(intRows,A290_EG2_E4_cls_desc))
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E5_ar_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
			strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group_cls(intRows,A290_EG2_E5_ar_dt))
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls(intRows,A290_EG2_E3_bal_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                 '11
		    strData = strData & Chr(11) & Chr(12)           
		Next
	
		Response.Write "<Script Language=VBScript> "															& vbCr  
		Response.Write " With parent "																			& vbCr 
		Response.Write " .ggoSpread.Source = .frm1.vspdData1"													& vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur1,.C_ARBALAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur1,.C_ARCLSAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur1,.C_ARAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData1," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur1,.C_ARDCAMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    .frm1.vspdData1.Redraw = True   "                      & vbCr
		Response.Write " End With "																				& vbCr
		Response.Write "</Script>"  																			& vbCr	
	End If    

	'////////////////////////////////////
	'		��Ÿ���� ���� 
	'////////////////////////////////////
	strData = "" 

	If Not Missing(EG3_export_group_dc) And Not IsEmpty(EG3_export_group_dc) Then	
		For IntRows = 0 To intCount1
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E1_seq))
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E2_acct_cd))
   		    strData = strData & Chr(11) & ""
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E2_acct_nm))
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E2_dr_cr_fg))
   		    strData = strData & Chr(11) & ""
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E2_doc_cur))
			strData = strData & Chr(11) & ""   		       		    
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group_dc(intRows,A290_EG3_E2_xch_rate),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group_dc(intRows,A290_EG3_E3_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group_dc(intRows,A290_EG3_E3_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   		    strData = strData & Chr(11) & ConvSPChars(EG3_export_group_dc(intRows,A290_EG3_E3_dc_desc))		
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
		    strData = strData & Chr(11) & Chr(12)           
		Next

		Response.Write "<Script Language=VBScript> "					       & vbCr  
		Response.Write " With parent "									       & vbCr 
		Response.Write " .ggoSpread.Source = .frm1.vspdData                  " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & strData   & """ ,""F""" & vbCr
		Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & LngMaxRow + 1 & "," & LngMaxRow + iLngRow - 1 & ",.C_DocCur,.C_ItemAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write  "    .frm1.vspdData.Redraw = True   "                      & vbCr
		Response.Write " End With "									 	       & vbCr
		Response.Write "</Script>"  										   & vbCr		
	End If    

	Response.Write "<Script Language=VBScript> "							   & vbCr  
	Response.Write " With parent "											   & vbCr 
	Response.Write " .frm1.txtRcptNo.value = """ & I1_a_rcpt_no			& """" & vbCr
	Response.Write " .frm1.horgChangeId.value = """ & txthOrgChangeId	& """" & vbCr	
	Response.Write " .DbQueryOk	"										       & vbCr
    Response.Write " End With "									 		       & vbCr
    Response.Write "</Script>"  										       & vbCr          	

%>	
	
