<%
option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 채권관리 
'*  3. Program ID           : A3105mb1
'*  4. Program Name         : 입금등록및 채권반제 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +B21011 (Manager)
'                             +B21019 (조회용)
'*  7. Modified date(First) : 2001/02/22
'*  8. Modified date(Last)  : 2002/12/18
'*  9. Modifier (First)     : Chang Sung Hee
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************



'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

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
'												2. 조건부 
'##########################################################################################################

													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd	
On Error Resume Next
Err.Clear 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then
	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim pAr0049																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim IntRows
Dim IntCols
Dim sList
Dim vbIntRet
Dim intCount
Dim IntCount1
Dim LngMaxRow
Dim LngMaxRow1
Dim StrNextKey
Dim lgStrPrevKey
Dim lgIntFlgMode
dim test

' Com+ Conv. 변수 선언 
Dim pvStrGlobalCollection 

Dim I1_a_rcpt
Dim E1_b_biz_area
Dim E2_b_minor
Dim E3_a_rcpt_item
Dim E4_a_gl
Dim E5_b_acct_dept
Dim E6_b_biz_partner
DIm E7_b_bank
Dim E8_a_rcpt
Dim E9_a_allc_rcpt
Dim EG1_export_group
Dim EG6_export_group_dc

Dim arrCount
Dim lgCurrency
Dim strData

' 첨자 선언 
Const A308_I1_a_rcpt_no = 0

'[CONVERSION INFORMATION]  View Name : export b_biz_area
Const A308_E1_biz_area_cd = 0
Const A308_E1_biz_area_nm = 1

'[CONVERSION INFORMATION]  View Name : export b_minor
Const A308_E2_minor_nm = 0

'[CONVERSION INFORMATION]  View Name : export a_rcpt_item
Const A308_E3_rcpt_type = 0
Const A308_E3_note_no = 1
Const A308_E3_bank_acct_no = 2

'[CONVERSION INFORMATION]  View Name : export a_gl
Const A308_E4_gl_no = 0

'[CONVERSION INFORMATION]  View Name : export b_acct_dept
Const A308_E5_dept_cd = 0
Const A308_E5_dept_nm = 1

'[CONVERSION INFORMATION]  View Name : export b_biz_partner
Const A308_E6_bp_cd = 0
Const A308_E6_bp_nm = 1

'[CONVERSION INFORMATION]  View Name : export b_bank
Const A308_E7_bank_cd = 0
Const A308_E7_bank_nm = 1

'[CONVERSION INFORMATION]  View Name : export a_rcpt
Const A308_E8_rcpt_no = 0
Const A308_E8_rcpt_dt = 1
Const A308_E8_ref_no = 2
Const A308_E8_doc_cur = 3
Const A308_E8_xch_rate = 4
Const A308_E8_rcpt_fg = 5
Const A308_E8_rcpt_amt = 6
Const A308_E8_rcpt_loc_amt = 7
Const A308_E8_allc_amt = 8
Const A308_E8_allc_loc_amt = 9
Const A308_E8_conf_fg = 10
Const A308_E8_temp_gl_no = 11
Const A308_E8_acct_cd = 12
Const A308_E8_acct_nm = 13
Const A308_E8_rcpt_desc = 14

'[CONVERSION INFORMATION]  View Name : export a_allc_rcpt
Const A308_E9_allc_no = 0
Const A308_E9_allc_dt = 1
Const A308_E9_allc_type = 2
Const A308_E9_dc_amt = 3
Const A308_E9_dc_loc_amt = 4

'[CONVERSION INFORMATION]  Group Name : export_group
'[CONVERSION INFORMATION]  View Name : export_cls b_biz_area
Const A308_EG1_E1_b_biz_area_biz_area_cd = 0
Const A308_EG1_E1_b_biz_area_biz_area_nm = 1
'[CONVERSION INFORMATION]  View Name : export_cls a_acct
Const A308_EG2_E1_a_acct_acct_cd = 2
Const A308_EG2_E1_a_acct_acct_nm = 3
'[CONVERSION INFORMATION]  View Name : export_cls b_acct_dept
Const A308_EG3_E1_b_acct_dept_dept_cd = 4
Const A308_EG3_E1_b_acct_dept_dept_nm = 5
'[CONVERSION INFORMATION]  View Name : export_cls a_cls_ar
Const A308_EG4_E1_a_cls_ar_cls_dt = 6
Const A308_EG4_E1_a_cls_ar_doc_cur = 7
Const A308_EG4_E1_a_cls_ar_cls_amt = 8
Const A308_EG4_E1_a_cls_ar_cls_loc_amt = 9
Const A308_EG4_E1_a_cls_ar_dc_amt = 10
Const A308_EG4_E1_a_cls_ar_dc_loc_amt = 11
Const A308_EG4_E1_a_cls_ar_cls_ar_desc = 12
'[CONVERSION INFORMATION]  View Name : export_cls a_open_ar
Const A308_EG5_E1_a_open_ar_ar_no = 13
Const A308_EG5_E1_a_open_ar_ar_dt = 14
Const A308_EG5_E1_a_open_ar_ar_due_dt = 15
Const A308_EG5_E1_a_open_ar_ar_amt = 16
Const A308_EG5_E1_a_open_ar_ar_loc_amt = 17
Const A308_EG5_E1_a_open_ar_cls_amt = 18
Const A308_EG5_E1_a_open_ar_cls_loc_amt = 19
Const A308_EG5_E1_a_open_ar_bal_amt = 20
Const A308_EG5_E1_a_open_ar_bal_loc_amt = 21
Const A308_EG5_E1_a_open_ar_ref_no = 22

'[CONVERSION INFORMATION]  Group Name : export_group_dc
'[CONVERSION INFORMATION]  View Name : export_dc a_acct
Const A308_EG6_E1_a_acct_acct_cd = 0
Const A308_EG6_E1_a_acct_acct_nm = 1
    
Const A308_EG7_E1_a_rcpt_dc_seq = 2
Const A308_EG7_E1_a_rcpt_dc_dc_amt = 3
Const A308_EG7_E1_a_rcpt_dc_dc_loc_amt = 4
Const A308_EG7_E1_a_rcpt_dc_dc_desc = 5

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")

'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################
	Set pAr0049 = Server.CreateObject("PARG025.cALkUpDirectRcSvr")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If
	
	LngMaxRow  = CLng(Request("txtMaxRows"))												'☜: Fetechd Count      
	LngMaxRow1  = CLng(Request("txtMaxRows1"))
	
    Redim I1_a_rcpt(A308_I1_a_rcpt_no+4)
    I1_a_rcpt(A308_I1_a_rcpt_no)   = Request("txtRcptNo")
	I1_a_rcpt(A308_I1_a_rcpt_no+1) = lgAuthBizAreaCd
	I1_a_rcpt(A308_I1_a_rcpt_no+2) = lgInternalCd
	I1_a_rcpt(A308_I1_a_rcpt_no+3) = lgSubInternalCd
	I1_a_rcpt(A308_I1_a_rcpt_no+4) = lgAuthUsrID	

	Call pAr0049.A_LOOKUP_DIRECT_RCPT_SVR(gStrGlobalCollection,I1_a_rcpt, E1_b_biz_area, E2_b_minor,E3_a_rcpt_item, E4_a_gl, E5_b_acct_dept, E6_b_biz_partner, E7_b_bank, E8_a_rcpt, E9_a_allc_rcpt, EG1_export_group, EG6_export_group_dc)

	'-----------------------
	'Com Action Area
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr0049 = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pAr0049 = Nothing
	
	lgCurrency = ConvSPChars(E8_a_rcpt(A308_E8_doc_cur))	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " With parent " & vbCr
	Response.Write " .frm1.hArDocCur.value  = """ & ConvSPChars(EG1_export_group(0,A308_EG4_E1_a_cls_ar_doc_cur)) & """" & vbCr
		
	Response.Write ".frm1.txtRcptDt.Text			= """ & UNIDateClientFormat(E8_a_rcpt(A308_E8_rcpt_dt))	& """" & vbCr
	Response.Write ".frm1.txtDeptCd.Value			= """ & ConvSPChars(E5_b_acct_dept(A308_E5_dept_cd))	& """" & vbCr
	Response.Write ".frm1.txtDeptNm.Value			= """ & ConvSPChars(E5_b_acct_dept(A308_E5_dept_nm))	& """" & vbCr
	Response.Write ".frm1.txtBankCd.Value			= """ & ConvSPChars(E7_b_bank(A308_E7_bank_cd))			& """" & vbCr
	Response.Write ".frm1.txtBankNm.Value		    = """ & ConvSPChars(E7_b_bank(A308_E7_bank_nm))			& """" & vbCr
    Response.Write ".frm1.txtBpCd.Value				= """ & ConvSPChars(E6_b_biz_partner(A308_E6_bp_cd))	& """" & vbCr
    Response.Write ".frm1.txtBpnM.Value				= """ & ConvSPChars(E6_b_biz_partner(A308_E6_bp_nm))	& """" & vbCr
    Response.Write ".frm1.txtBankAcct.Value			= """ & ConvSPChars(E3_a_rcpt_item(A308_E3_bank_acct_no)) & """" & vbCr
    Response.Write ".frm1.txtInputType.Value		= """ & ConvSPChars(E3_a_rcpt_item(A308_E3_rcpt_type))	& """" & vbCr
    Response.Write ".frm1.txtInputTypeNm.Value		= """ & ConvSPChars(E2_b_minor(A308_E2_minor_nm))		& """" & vbCr
    Response.Write ".frm1.txtCheckCD.Value			= """ & ConvSPChars(E3_a_rcpt_item(A308_E3_note_no))	& """" & vbCr
    Response.Write ".frm1.txtDocCur.value			= """ & ConvSPChars(E8_a_rcpt(A308_E8_doc_cur))			& """" & vbCr
    Response.Write ".frm1.txtXchRate.Text			= """ & ConvSPChars(E8_a_rcpt(A308_E8_xch_rate))		& """" & vbCr
    Response.Write ".frm1.txtTempGlNo.value			= """ & ConvSPChars(E8_a_rcpt(A308_E8_temp_gl_no))		& """" & vbCr
	Response.Write ".frm1.txtGlNo.value				= """ & ConvSPChars(E4_a_gl(A308_E4_gl_no))				& """" & vbCr
	Response.Write ".frm1.txtRcptDesc.value			= """ & ConvSPChars(E8_a_rcpt(A308_E8_rcpt_desc))		& """" & vbCr
	Response.Write ".frm1.txtAcctCd.value			= """ & ConvSPChars(E8_a_rcpt(A308_E8_acct_cd))			& """" & vbCr
	Response.Write ".frm1.txtAcctNm.value			= """ & ConvSPChars(E8_a_rcpt(A308_E8_acct_nm))			& """" & vbCr			

	Response.Write ".frm1.txtRcptAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E8_a_rcpt(A308_E8_rcpt_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbCr
	Response.Write ".frm1.txtRcptLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E8_a_rcpt(A308_E8_rcpt_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".frm1.txtDcAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_rcpt(A308_E9_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbCr
	Response.Write ".frm1.txtDcLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_rcpt(A308_E9_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".frm1.txtAllcNo.Value			= """ & ConvSPChars(E9_a_allc_rcpt(A308_E9_allc_no))	& """" & vbCr
	Response.Write " End With "                 & vbCr
    Response.Write "</Script>"	

    
    intCount = UBound(EG1_export_group,1)
	StrNextKey = ""   ' import view
	
	lgCurrency = ConvSPChars(EG1_export_group(0,A308_EG4_E1_a_cls_ar_doc_cur))

	If IsEmpty(EG1_export_group) = False and IsArray(EG1_export_group) = True Then    
		For IntRows = 0 To intCount		
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_ar_no))
		    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_ar_dt))    
		    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_ar_due_dt))
			strData = strData & Chr(11) & UCase(ConvSPChars(EG1_export_group(IntRows,A308_EG4_E1_a_cls_ar_doc_cur)))
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_ar_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG5_E1_a_open_ar_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG4_E1_a_cls_ar_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A308_EG4_E1_a_cls_ar_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG4_E1_a_cls_ar_cls_ar_desc))
		    
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG2_E1_a_acct_acct_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG2_E1_a_acct_acct_nm))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG1_E1_b_biz_area_biz_area_cd))
		    strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A308_EG1_E1_b_biz_area_biz_area_nm))
		    strData = strData & Chr(11) & LngMaxRow + IntRows + 1
			strData = strData & Chr(11) & Chr(12)                                    
		Next
	End If

    Response.Write "<Script Language=VBScript> "                                                          & vbCr  
    Response.Write " With parent "                                                                        & vbCr 
    Response.Write " .ggoSpread.Source          = .frm1.vspdData1 "								          & vbCr
    Response.Write " .ggoSpread.SSShowData        """ & strData											& """" & vbCr
    Response.Write " .lgStrPrevKey				= """ & StrNextKey										& """" & vbCr
    Response.Write " End With "                                                                           & vbCr
    Response.Write "</Script>"  																		  & vbCr	
	
	strData = "" 

	If IsEmpty(EG6_export_group_dc) = False and IsArray(EG6_export_group_dc) = True Then    
		intCount1 = UBound(EG6_export_group_dc,1)
		For IntRows = 0 To intCount1
    	    strData = strData & Chr(11) & ConvSPChars(EG6_export_group_dc(IntRows,A308_EG7_E1_a_rcpt_dc_seq))
    	    strData = strData & Chr(11) & ConvSPChars(EG6_export_group_dc(IntRows,A308_EG6_E1_a_acct_acct_cd))
    	    strData = strData & Chr(11) & ""
    	    strData = strData & Chr(11) & ConvSPChars(EG6_export_group_dc(IntRows,A308_EG6_E1_a_acct_acct_nm))

    	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG6_export_group_dc(IntRows,A308_EG7_E1_a_rcpt_dc_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
    	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG6_export_group_dc(IntRows,A308_EG7_E1_a_rcpt_dc_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")

            strData = strData & Chr(11) & LngMaxRow + IntRows + 1                              '11
            strData = strData & Chr(11) & Chr(12)           
		Next
	End If

    Response.Write "<Script Language=VBScript> "														   & vbCr  
    Response.Write " With parent "																		   & vbCr 
	Response.Write " .ggoSpread.Source    = .frm1.vspdData "											   & vbCr
	Response.Write " .ggoSpread.SSShowData  """ & strData											& """" & vbCr
	Response.Write " .DbQueryOK "																		   & vbCr
	Response.Write ".frm1.txtAcctCd.value = """ & ConvSPChars(E8_a_rcpt(A308_E8_acct_cd))			& """" & vbCr
	Response.Write ".frm1.txtAcctNm.value = """ & ConvSPChars(E8_a_rcpt(A308_E8_acct_nm))			& """" & vbCr
	
    Response.Write " End With "                                                                            & vbCr
    Response.Write "</Script>"  																		   & vbCr          	

%>		
