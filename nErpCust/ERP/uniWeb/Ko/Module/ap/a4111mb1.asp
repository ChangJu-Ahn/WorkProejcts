<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account 
'*  2. Function Name        : 
'*  3. Program ID           : a4111mb1.adp
'*  4. Program Name         : 채무/채권 상계 조회 Logic	
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : YOU SO EUN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
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
Err.Clear 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

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
ElseIf Request("txtClearNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
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
Dim iPAPG055																	'☆ : 조회용 ComPlus Dll 사용 변수 
Dim IntRows
Dim strData
Dim intCount
Dim IntCount1
Dim LngMaxRow
Dim lgCurrency

Dim I1_a_clear_ap_ar 
Dim E1_a_clear_ap_ar 
Dim E2_a_clear_ap_ar_etc
Dim EG1_export_group_cls_ar 
Dim EG2_export_group_cls_ap 

Const A314_E1_clear_no = 0
Const A314_E1_clear_dt = 1
Const A314_E1_clear_gl_no1 = 2
Const A314_E1_clear_temp_gl_no1 = 3
Const A314_E1_clear_org_change_id1 = 4
Const A314_E1_clear_dept_cd = 5
Const A314_E1_clear_dept_nm = 6
Const A314_E1_clear_desc = 7

Const A314_E2_bp_cd = 0    
Const A314_E2_bp_nm = 1
Const A314_E2_doc_cur = 2

Const A314_EG1_E4_ar_no = 0
Const A314_EG1_E3_acct_cd = 1
Const A314_EG1_E3_acct_nm = 2
Const A314_EG1_E4_ar_dt = 3
Const A314_EG1_E5_ar_due_dt = 4
Const A314_EG1_E4_ar_amt = 5
Const A314_EG1_E4_bal_amt = 6
Const A314_EG1_E4_cls_amt = 7
Const A314_EG1_E4_cls_loc_amt = 8
Const A314_EG1_E5_cls_ar_desc = 9
    
Const A314_EG2_E4_ap_no = 0
Const A314_EG2_E3_acct_cd = 1
Const A314_EG2_E3_acct_nm = 2
Const A314_EG2_E4_ap_dt = 3
Const A314_EG2_E4_ap_due_dt = 4
Const A314_EG2_E4_ap_amt = 5
Const A314_EG2_E4_bal_amt = 6
Const A314_EG2_E5_cls_amt = 7
Const A314_EG2_E5_cls_loc_amt = 8
Const A314_EG2_E5_cls_ap_desc = 9

'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################

'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Const A314_I1_clear_no = 0

	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

    Redim I1_a_clear_ap_ar(A314_I1_clear_no+4)
    I1_a_clear_ap_ar(A314_I1_clear_no)   = Trim(Request("txtClearNo"))
	I1_a_clear_ap_ar(A314_I1_clear_no+1) = lgAuthBizAreaCd
	I1_a_clear_ap_ar(A314_I1_clear_no+2) = lgInternalCd
	I1_a_clear_ap_ar(A314_I1_clear_no+3) = lgSubInternalCd
	I1_a_clear_ap_ar(A314_I1_clear_no+4) = lgAuthUsrID	
	
	'-----------------------------------------
	'Data manipulate  area(import view match)
	'-----------------------------------------
'	I1_a_clear_ap_ar = Trim(Request("txtClearNo"))

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set iPAPG055 = Server.CreateObject("PAPG055.cALkUpClearApArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Call iPAPG055.A_LOOKUP_CLEAR_AP_AR_SVR(gStrGlobalCollection, I1_a_clear_ap_ar, E1_a_clear_ap_ar, E2_a_clear_ap_ar_etc, _
	                                    EG1_export_group_cls_ar, EG2_export_group_cls_ap)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG055 = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Set iPAPG055 = Nothing

	lgCurrency = ConvSPChars(E2_a_clear_ap_ar_etc(A314_E2_doc_cur))

	Response.Write "<Script Language=vbscript>"																		& vbCr
	Response.Write " With parent.frm1 "																				& vbCr
	Response.Write " .txtDeptCd.Value    = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_dept_cd))		 & """" & vbCr
	Response.Write " .txtDeptNm.Value    = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_dept_nm))		 & """" & vbCr
	Response.Write " .txtAllcDt.text     = """ & UNIDateClientFormat(E1_a_clear_ap_ar(A314_E1_clear_dt))	 & """" & vbCr
	Response.Write " .txtDesc.Value		 = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_desc))			 & """" & vbCr	
	Response.Write " .txtGlNo.Value      = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_gl_no1))		 & """" & vbCr
	Response.Write " .txtTempGlNo.Value  = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_temp_gl_no1))	 & """" & vbCr	
	Response.Write " .hOrgChangeId.Value = """ & ConvSPChars(E1_a_clear_ap_ar(A314_E1_clear_org_change_id1)) & """" & vbCr	
	Response.Write " .txtBpCd.Value      = """ & ConvSPChars(E2_a_clear_ap_ar_etc(A314_E2_bp_cd))			 & """" & vbCr
	Response.Write " .txtBpNm.Value      = """ & ConvSPChars(E2_a_clear_ap_ar_etc(A314_E2_bp_nm))			 & """" & vbCr
	Response.Write " .txtDocCur.Value    = """ & ConvSPChars(E2_a_clear_ap_ar_etc(A314_E2_doc_cur))			 & """" & vbCr
	Response.Write " End With "																						& vbCr
    Response.Write "</Script>"																						& vbCr
    
    intCount = UBound(EG2_export_group_cls_ap,1)
    intCount1 = UBound(EG1_export_group_cls_ar,1)

	strData = "" 

	For IntRows = 0 To intCount
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls_ap(IntRows,A314_EG2_E4_ap_no))
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls_ap(IntRows,A314_EG2_E3_acct_cd))    	    
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls_ap(IntRows,A314_EG2_E3_acct_nm))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group_cls_ap(IntRows,A314_EG2_E4_ap_dt))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group_cls_ap(IntRows,A314_EG2_E4_ap_due_dt))
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls_ap(IntRows,A314_EG2_E4_ap_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls_ap(IntRows,A314_EG2_E4_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls_ap(IntRows,A314_EG2_E5_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_cls_ap(IntRows,A314_EG2_E5_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_cls_ap(IntRows,A314_EG2_E5_cls_ap_desc))   	    
        strData = strData & Chr(11) & LngMaxRow + IntRows                                 
        strData = strData & Chr(11) & Chr(12)                    	
	Next

	Response.Write "<Script Language=vbscript>								" & vbCr
	Response.Write " With parent											" & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData                     " & vbCr
	Response.Write " .frm1.vspdData.ReDraw = False							" & vbCr	
	Response.Write " .ggoSpread.SSShowData """ & strData & """,""F"""         & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ApAmt   ,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ApRemAmt,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ApClsAmt,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " End With												" & vbCr
    Response.Write "</Script>												" & vbCr

	strData = ""

	For IntRows = 0 To intCount1
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_ar_no))
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_cls_ar(IntRows,A314_EG1_E3_acct_cd))    	    
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_cls_ar(IntRows,A314_EG1_E3_acct_nm))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_ar_dt))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group_cls_ar(IntRows,A314_EG1_E5_ar_due_dt))
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_ar_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
   	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_cls_ar(IntRows,A314_EG1_E4_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_cls_ar(IntRows,A314_EG1_E5_cls_ar_desc))
        strData = strData & Chr(11) & LngMaxRow + IntRows
        strData = strData & Chr(11) & Chr(12)
	Next

	Response.Write "<Script Language=vbscript>								" & vbCr
	Response.Write " With parent											" & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData1			  	    " & vbCr
	Response.Write " .frm1.vspdData1.ReDraw = False							" & vbCr		
	Response.Write " .ggoSpread.SSShowData """ & strData & """,""F"""         & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ArAmt   ,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ArRemAmt,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & -1 & "," & -1 & ",.frm1.txtDocCur.value,.C_ArClsAmt,""A"",""I"",""X"",""X"")" & vbCr
	Response.Write " .frm1.htxtClearNo.value = (""" & I1_a_clear_ap_ar  & """)" & vbCr
	Response.Write " .frm1.vspdData.ReDraw = True							" & vbCr		
	Response.Write " .frm1.vspdData1.ReDraw = True							" & vbCr
	Response.Write " .DbQueryOk												" & vbCr
	Response.Write " End With												" & vbCr
    Response.Write "</Script>												" & vbCr	

%>    
