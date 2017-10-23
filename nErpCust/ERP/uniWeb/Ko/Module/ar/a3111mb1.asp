<%
'**********************************************************************************************
'*  1. Module Name          : 입금반제 
'*  2. Function Name        : 
'*  3. Program ID           : a3111mb1.aps
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Comproxy List        : +Ar0049r
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/06/17
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : Chang Sung Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
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
	Call DisplayMsgBox("700118", vbOKOnly, "", "", I_MKSCRIPT)	'조회요청만 할 수 있습니다.
	Response.End
	Call HideStatusWnd		 
ElseIf Request("txtAllcNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다!
	Response.End
	Call HideStatusWnd		 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim pAr0049r																'☆ : 조회용 ComProxy Dll 사용 변수 
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

Dim I1_a_rcpt_dc 
Dim I2_a_open_ar 
Dim I3_a_rcpt 
Dim I4_a_allc_rcpt 
Dim E1_b_biz_area 
Dim E2_a_rcpt_dc 
Dim E3_a_open_ar 
Dim E4_a_rcpt 
Dim E5_a_rcpt 
Dim E6_b_acct_dept 
Dim E7_b_biz_partner 
Dim E8_a_allc_rcpt 
Dim E9_a_gl 
Dim EG1_export_group_assn 
Dim EG2_export_group 
Dim EG3_export_group_dc 

Const A296_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_area
Const A296_E1_biz_area_nm = 1

Const A296_E5_dept_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_acct_dept
Const A296_E5_dept_nm = 1

Const A296_E6_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_partner
Const A296_E6_bp_nm = 1

Const A296_E8_allc_no = 0    '[CONVERSION INFORMATION]  View Name : export a_allc_rcpt
Const A296_E8_allc_dt = 1
Const A296_E8_allc_type = 2
Const A296_E8_ref_no = 3
Const A296_E8_allc_amt = 4
Const A296_E8_allc_loc_amt = 5
Const A296_E8_dc_amt = 6
Const A296_E8_dc_loc_amt = 7
Const A296_E8_allc_rcpt_desc = 8
Const A296_E8_temp_gl_no = 9

Const A296_EG1_E1_acct_cd = 0    '[CONVERSION INFORMATION]  View Name : export_dc a_acct
Const A296_EG1_E1_acct_nm = 1
Const A296_EG1_E2_seq = 2    '[CONVERSION INFORMATION]  View Name : export a_rcpt_dc
Const A296_EG1_E2_dc_amt = 3
Const A296_EG1_E2_dc_loc_amt = 4

Const A296_EG2_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export_rcpt b_biz_area
Const A296_EG2_E1_biz_area_nm = 1
Const A296_EG2_E2_dept_cd = 2    '[CONVERSION INFORMATION]  View Name : export_rcpt b_acct_dept
Const A296_EG2_E2_dept_nm = 3
Const A296_EG2_E3_acct_cd = 4    '[CONVERSION INFORMATION]  View Name : export a_acct
Const A296_EG2_E3_acct_nm = 5
Const A296_EG2_E4_allc_dt = 6    '[CONVERSION INFORMATION]  View Name : export a_allc_rcpt_assn
Const A296_EG2_E4_allc_amt = 7
Const A296_EG2_E4_allc_loc_amt = 8
Const A296_EG2_E5_rcpt_no = 9    '[CONVERSION INFORMATION]  View Name : export a_rcpt
Const A296_EG2_E5_rcpt_dt = 10
Const A296_EG2_E5_rcpt_amt = 11
Const A296_EG2_E5_rcpt_loc_amt = 12
Const A296_EG2_E5_allc_amt = 13
Const A296_EG2_E5_allc_loc_amt = 14
Const A296_EG2_E5_bal_amt = 15
Const A296_EG2_E5_bal_loc_amt = 16
Const A296_EG2_E5_allc_assn_desc = 17    

Const A296_EG3_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export_cls_ar b_biz_area
Const A296_EG3_E1_biz_area_nm = 1
Const A296_EG3_E2_dept_cd = 2    '[CONVERSION INFORMATION]  View Name : export_cls b_acct_dept
Const A296_EG3_E2_dept_nm = 3
Const A296_EG3_E3_cls_dt = 4    '[CONVERSION INFORMATION]  View Name : export a_cls_ar
Const A296_EG3_E3_ar_due_dt = 5
Const A296_EG3_E3_cls_amt = 6
Const A296_EG3_E3_cls_loc_amt = 7
Const A296_EG3_E3_dc_amt = 8
Const A296_EG3_E3_dc_loc_amt = 9
Const A296_EG3_E3_cls_ar_no = 10
Const A296_EG3_E4_acct_cd = 11    '[CONVERSION INFORMATION]  View Name : export_cls_ar a_acct
Const A296_EG3_E4_acct_nm = 12
Const A296_EG3_E5_ar_no = 13    '[CONVERSION INFORMATION]  View Name : export a_open_ar
Const A296_EG3_E5_ar_dt = 14
Const A296_EG3_E5_ar_amt = 15
Const A296_EG3_E5_ar_loc_amt = 16
Const A296_EG3_E5_cls_amt = 17
Const A296_EG3_E5_cls_loc_amt = 18
Const A296_EG3_E5_bal_amt = 19
Const A296_EG3_E5_bal_loc_amt = 20
Const A296_EG3_E5_ref_no = 21
Const A296_EG3_E5_cls_ar_desc = 22
Const A296_EG3_E5_ar_doc_cur = 23

Const A296_E10_cost_cd = 0    '[CONVERSION INFORMATION]  View Name : xxx_export b_cost_center
Const A296_E10_cost_nm = 1

	I4_a_allc_rcpt = Trim(Request("txtAllcNo"))
	I2_a_open_ar = Trim(lgStrPrevKey1)
	I3_a_rcpt = Trim(lgStrPrevKey)
	I1_a_rcpt_dc = Trim(lgStrPrevKeyDtl)

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set pAr0049r = Server.CreateObject("PARG055.cALkUpAllcRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Call pAr0049r.A_LOOKUP_ALLC_RCPT_SVR(gStrGlobalCollection, I1_a_rcpt_dc, I2_a_open_ar,I3_a_rcpt, I4_a_allc_rcpt, E1_b_biz_area, E2_a_rcpt_dc, E3_a_open_ar, E4_a_rcpt, E5_b_acct_dept, E6_b_biz_partner, E7_a_rcpt, E8_a_allc_rcpt, E9_a_gl, EG1_export_group_dc, EG2_export_group_assn, EG3_export_group)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr0049r = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If
		
	Set pAr0049r = Nothing

	lgCurrency = ConvSPChars(E7_a_rcpt)

	Response.Write "<Script Language=vbscript>"																		& vbCr
	Response.Write " With parent.frm1 "																				& vbCr
	Response.Write " .hArDocCur.value  = """ & ConvSPChars(EG3_export_group(0,A296_EG3_E5_ar_doc_cur))		 & """" & vbCr
	
	Response.Write ".htxtAllcNo.value				= """ & ConvSPChars(I4_a_allc_rcpt)								& """" & vbCr
	Response.Write ".txtBpCd.Value					= """ & ConvSPChars(E6_b_biz_partner(A296_E6_bp_cd))			& """" & vbCr
	Response.Write ".txtBpNm.Value					= """ & ConvSPChars(E6_b_biz_partner(A296_E6_bp_cd))			& """" & vbCr
	Response.Write ".txtDeptCd.Value				= """ & ConvSPChars(E5_b_acct_dept(A296_E5_dept_cd))			& """" & vbCr
	Response.Write ".txtDeptNm.Value				= """ & ConvSPChars(E5_b_acct_dept(A296_E5_dept_nm))			& """" & vbCr
	Response.Write ".txtDocCur.Value				= """ & ConvSPChars(E7_a_rcpt)									& """" & vbCr
	Response.Write ".txtAllcDt.text					= """ & UNIDateClientFormat(E8_a_allc_rcpt(A296_E8_allc_dt))	& """" & vbCr
	Response.Write ".txtTempGlNo.value				= """ & ConvSPChars(E8_a_allc_rcpt(A296_E8_temp_gl_no))			& """" & vbCr
	Response.Write ".txtGlNo.value					= """ & ConvSPChars(E9_a_gl)									& """" & vbCr
	Response.Write ".txtRcptDesc.value				= """ & ConvSPChars(E8_a_allc_rcpt(A296_E8_allc_rcpt_desc))		& """" & vbCr	
	
	Response.Write " End With "																		                & vbCr
    Response.Write "</Script>"																						& vbCr
    
    
    intCount = UBound(EG2_export_group_assn,1)
    intCount0 = UBound(EG3_export_group,1)
    IntCount1 = UBound(EG1_export_group_dc,1)

    If IntCount1 = "" or IntCount1 = null Then
		IntCount1 = -1    
	End IF
    
    If Trim(I3_a_rcpt) = EG2_export_group_assn(intCount, A296_EG2_E5_rcpt_no) OR _
       Trim(I2_a_open_ar) = EG3_export_group(intCount0, A296_EG3_E5_ar_no) Then 
		I2_a_open_ar = "" 
		I3_a_rcpt = "" 
    Else
		I2_a_open_ar = E3_a_open_ar 
		I3_a_rcpt = E4_a_rcpt 		
    End If

	For IntRows = 0 To intCount
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E5_rcpt_no))
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E3_acct_cd))
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E3_acct_nm))
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E1_biz_area_cd))
   	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E1_biz_area_nm))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group_assn(IntRows,A296_EG2_E5_rcpt_dt))
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_assn(IntRows,A296_EG2_E5_rcpt_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_assn(IntRows,A296_EG2_E5_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_assn(IntRows,A296_EG2_E5_allc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_assn(IntRows,A296_EG2_E5_allc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group_assn(IntRows,A296_EG2_E5_allc_assn_desc))
        strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
        strData = strData & Chr(11) & Chr(12)   
	Next

    Response.Write "<Script Language=VBScript> "															& vbCr  
    Response.Write " With parent "																			& vbCr 
	Response.Write " .ggoSpread.Source = .frm1.vspdData1 "													& vbCr
	Response.Write " .ggoSpread.SSShowData """ & strData 													& """" & vbCr
    Response.Write " End With "																				& vbCr
    Response.Write "</Script>"  																			& vbCr	
		
	strData = "" 

	lgCurrency = ConvSPChars(EG3_export_group(0,A296_EG3_E5_ar_doc_cur))

	For IntRows = 0 To intCount0
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E5_ar_no))
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E4_acct_cd))    	    
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E4_acct_nm))
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E1_biz_area_cd))    	    
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E1_biz_area_nm))	    	    
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG3_export_group(intRows,A296_EG3_E5_ar_dt))
   	    strData = strData & Chr(11) & UNIDateClientFormat(EG3_export_group(intRows,A296_EG3_E3_ar_due_dt))
   	    strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E5_ar_doc_cur))	    	    
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E5_ar_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E5_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E5_cls_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E5_cls_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E3_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG3_export_group(intRows,A296_EG3_E3_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		strData = strData & Chr(11) & ConvSPChars(EG3_export_group(intRows,A296_EG3_E5_cls_ar_desc))
        strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                 '11
        strData = strData & Chr(11) & Chr(12)           
	Next

	
    Response.Write "<Script Language=VBScript> "															& vbCr  
    Response.Write " With parent "																			& vbCr 
    Response.Write " .ggoSpread.Source = .frm1.vspdData10 "													& vbCr
    Response.Write " .ggoSpread.SSShowData """ & strData 													& """" & vbCr
    Response.Write " End With "																				& vbCr
    Response.Write "</Script>"  																			& vbCr	
    
	strData = "" 

	For IntRows = 0 To intCount1
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_dc(intRows,A296_EG1_E2_seq))
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_dc(intRows,A296_EG1_E1_acct_cd))
   	    strData = strData & Chr(11) & ""
   	    strData = strData & Chr(11) & ConvSPChars(EG1_export_group_dc(intRows,A296_EG1_E1_acct_nm))
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_dc(intRows,A296_EG1_E2_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group_dc(intRows,A296_EG1_E2_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
        strData = strData & Chr(11) & LngMaxRow + IntRows + 1                                  '11
        strData = strData & Chr(11) & Chr(12)           
	Next

    Response.Write "<Script Language=VBScript> "															& vbCr  
    Response.Write " With parent "																			& vbCr 
    Response.Write " .ggoSpread.Source = .frm1.vspdData "													& vbCr
    Response.Write " .ggoSpread.SSShowData		  """ & strData 											& """" & vbCr
    
    Response.Write " .lgStrPrevKey				= """ & I3_a_rcpt											& """" & vbCr
    Response.Write " .lgStrPrevKey1				= """ & I2_a_open_ar										& """" & vbCr
    Response.Write " .lgStrPrevKeyDtl			= """ & I1_a_rcpt_dc										& """" & vbCr

	Response.Write " .frm1.htxtAllcNo.value		= """ & I4_a_allc_rcpt										& """" & vbCr
	Response.Write " .DbQueryOk	"																			& vbCr
    Response.Write " End With "																				& vbCr
    Response.Write "</Script>"  																			& vbCr          	

%>	
	
