<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : Multi Alloction Query
'*  3. Program ID        : a3125mb1
'*  4. Program 이름      : 멀티입금(조회)
'*  5. Program 설명      : 멀티입금 조회 
'*  6. Complus 리스트    : PARG060
'*  7. 최초 작성년월일   : 2003/03/25
'*  8. 최종 수정년월일   : 2003/03/25
'*  9. 최초 작성자       : 정용균 
'* 10. 최종 작성자       : 정용균 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************

														'☜ : ASP가 캐쉬되지 않도록 한다.
														'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2.- 조건부 
'##########################################################################################################
																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

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

 
strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then'
	Response.End
	Call HideStatusWnd		 
ElseIf strMode <> CStr(UID_M0001) Then										'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Call DisplayMsgBox("700118", vbOKOnly, "", "", I_MKSCRIPT)				'조회요청만 할 수 있습니다.
	Response.End
	Call HideStatusWnd		 
ElseIf Trim(Request("txtRcptNo")) = "" Then									'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)				'조회 조건값이 비어있습니다!
	Response.End
	Call HideStatusWnd		 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim iPARG060																'☆ : 조회용 ComPlus Dll 사용 변수 
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
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Call iPARG060.A_LOOKUP_ALLC_RCPT_MULTI_SVR(gStrGlobalCollection, I1_a_rcpt_no, E1_a_allc_rcpt,E2_a_sum_amt, _
	                                      EG1_export_group_allc, EG2_export_group_cls, EG3_export_group_dc)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG060 = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If
		
	Set iPARG060 = Nothing

	'//////////////////////////////////////////////////////////////////
	'  멀티입금 헤더 정보 
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
	'		멀티반제내역 정보 
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
	'		채권반제내역 정보 
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
	'		기타내역 정보 
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
	
