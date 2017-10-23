<%
'**********************************************************************************************
'*  1. Module Name          : 선수금반제 
'*  2. Function Name        : 
'*  3. Program ID           : a3108mb1.aps
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Comproxy List        : +Ar0049pr
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

Call HideStatusWnd															'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then											'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Call ServerMesgBox("조회 요청만 할 수 있습니다!", vbInformation, I_MKSCRIPT)	'⊙: 조회 전용인데 다른 상태로 요청이 왔을 경우, 필요없으면 빼도 됨, 메세지는 ID값으로 사용해야 함 
	Response.End 
ElseIf Request("txtAllcNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)						'⊙:
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
Dim strData1
Dim strData2
Dim vbIntRet
Dim intCount
Dim IntCount1
'Dim IntCurSeq
Dim LngMaxRow
Dim StrNextKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgIntFlgMode
Dim test

Dim lgCurrency

' Com+ Conv. 변수 선언 
Dim pvStrGlobalCollection 

Dim I1_a_rcpt_dc 
Dim I2_a_open_ar 
Dim I3_a_allc_rcpt 
Dim E1_b_biz_area 
Dim E2_b_biz_partner 
Dim E3_a_rcpt_dc 
Dim E4_a_open_ar 
Dim EG1_export_group 
Dim EG2_export_group_dc 
Dim E5_a_gl 
Dim E6_b_acct_dept 
Dim E7_a_allc_rcpt 
Dim E8_f_prrcpt

Dim arrCount

' 첨자 선언 
Const A295_I3_a_allc_rcpt_rcpt_no = 0

Const A295_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_area
Const A295_E1_biz_area_nm = 1

Const A295_E2_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_partner
Const A295_E2_bp_nm = 1

Const A295_E4_prrcpt_no = 0    '[CONVERSION INFORMATION]  View Name : export f_prrcpt
Const A295_E4_prrcpt_dt = 1
Const A295_E4_doc_cur = 2
Const A295_E4_xch_rate = 3
Const A295_E4_prrcpt_amt = 4
Const A295_E4_loc_prrcpt_amt = 5
Const A295_E4_bal_amt = 6
Const A295_E4_loc_bal_amt = 7

Const A295_E5_allc_no = 0    '[CONVERSION INFORMATION]  View Name : export a_allc_rcpt
Const A295_E5_allc_dt = 1
Const A295_E5_allc_type = 2
Const A295_E5_ref_no = 3
Const A295_E5_allc_amt = 4
Const A295_E5_allc_loc_amt = 5
Const A295_E5_dc_amt = 6
Const A295_E5_dc_loc_amt = 7
Const A295_E5_allc_rcpt_desc = 8
Const A295_E5_temp_gl_no = 9

Const A295_E6_dept_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_acct_dept
Const A295_E6_dept_nm = 1

    '[CONVERSION INFORMATION]  Group Name : export_group_dc
Const A295_EG1_E1_acct_cd = 0    '[CONVERSION INFORMATION]  View Name : export_dc a_acct
Const A295_EG1_E1_acct_nm = 1
Const A295_EG1_E2_seq = 2    '[CONVERSION INFORMATION]  View Name : export a_rcpt_dc
Const A295_EG1_E2_dc_amt = 3
Const A295_EG1_E2_dc_loc_amt = 4

    '[CONVERSION INFORMATION]  Group Name : export_group
Const A295_EG2_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export_ar b_biz_area
Const A295_EG2_E1_biz_area_nm = 1
Const A295_EG2_E2_dept_cd = 2    '[CONVERSION INFORMATION]  View Name : export_cls b_acct_dept
Const A295_EG2_E2_dept_nm = 3
Const A295_EG2_E3_cls_dt = 4    '[CONVERSION INFORMATION]  View Name : export a_cls_ar
Const A295_EG2_E3_cls_amt = 5
Const A295_EG2_E3_cls_loc_amt = 6
Const A295_EG2_E3_dc_amt = 7
Const A295_EG2_E3_dc_loc_amt = 8
Const A295_EG2_E3_cls_ar_desc = 9
Const A295_EG2_E3_cls_ar_no = 10
Const A295_EG2_E4_acct_cd = 11    '[CONVERSION INFORMATION]  View Name : export_cls_ar a_acct
Const A295_EG2_E4_acct_nm = 12
Const A295_EG2_E5_ar_no = 13    '[CONVERSION INFORMATION]  View Name : export a_open_ar
Const A295_EG2_E5_ar_dt = 14
Const A295_EG2_E5_ar_amt = 15
Const A295_EG2_E5_ar_loc_amt = 16
Const A295_EG2_E5_cls_amt = 17
Const A295_EG2_E5_cls_loc_amt = 18
Const A295_EG2_E5_ar_due_dt = 19
Const A295_EG2_E5_bal_amt = 20
Const A295_EG2_E5_bal_loc_amt = 21
Const A295_EG2_E5_inv_doc_no = 22
Const A295_EG2_E5_ref_no = 23
Const A295_EG2_E5_doc_cur = 24

'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")
	lgStrPrevKey1 = Request("lgStrPrevKey1")

'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################

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

    Redim I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no+4)
    I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no)   = Trim(Request("txtAllcNo"))
	I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no+1) = lgAuthBizAreaCd
	I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no+2) = lgInternalCd
	I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no+3) = lgSubInternalCd
	I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no+4) = lgAuthUsrID	

	I1_a_rcpt_dc = lgStrPrevKey1
	I2_a_open_ar = lgStrPrevKey

	Set pAr0049 = Server.CreateObject("PARG040.cALkUpAllcPrSvr")

	'--------------------------------------------
	'Com action result check area(OS,internal)
	'--------------------------------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If
	
	'-------------------------------------------
	'Data manipulate  area(import view match)
	'-------------------------------------------
	Call pAr0049.A_LOOKUP_ALLC_PRERCPT_SVR(gStrGlobalCollection,I1_a_rcpt_dc,I2_a_open_ar,I3_a_allc_rcpt,E1_b_biz_area,E2_b_biz_partner, _
			E3_a_rcpt_dc,E4_a_open_ar,EG1_export_group,EG2_export_group_dc,E5_a_gl,E6_b_acct_dept,E7_a_allc_rcpt,E8_f_prrcpt)

	'-----------------------
	'Com Action Area
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr0049 = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pAr0049 = Nothing

	lgCurrency = ConvSPChars(E8_f_prrcpt(A295_E4_doc_cur))
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " With parent " & vbCr
		
	Response.Write " .frm1.htxtAllcNo.value			= """ & ConvSPChars(I3_a_allc_rcpt(A295_I3_a_allc_rcpt_rcpt_no)) & """" & vbCr
	Response.Write " .frm1.txtPrDt.text				= """ & UNIDateClientFormat(E8_f_prrcpt(A295_E4_prrcpt_dt))		& """" & vbCr
	Response.Write " .frm1.txtPrNo.Value			= """ & ConvSPChars(E8_f_prrcpt(A295_E4_prrcpt_no))				& """" & vbCr
	Response.Write " .frm1.txtAllcDt.text			= """ & UNIDateClientFormat(E7_a_allc_rcpt(A295_E5_allc_dt))	& """" & vbCr
	Response.Write " .frm1.txtDeptCd.Value			= """ & ConvSPChars(E6_b_acct_dept(A295_E6_dept_cd))			& """" & vbCr
	Response.Write " .frm1.txtDeptNm.Value			= """ & ConvSPChars(E6_b_acct_dept(A295_E6_dept_nm))			& """" & vbCr
	Response.Write " .frm1.txtBpCd.value			= """ & ConvSPChars(E2_b_biz_partner(A295_E2_bp_cd))			& """" & vbCr
	Response.Write " .frm1.txtBpNm.value			= """ & ConvSPChars(E2_b_biz_partner(A295_E2_bp_nm))			& """" & vbCr
	Response.Write " .frm1.txtDocCur.value			= """ & ConvSPChars(E8_f_prrcpt(A295_E4_doc_cur))				& """" & vbCr
	Response.Write " .frm1.txtTempGlNo.value		= """ & ConvSPChars(E7_a_allc_rcpt(A295_E5_temp_gl_no))			& """" & vbCr
	Response.Write " .frm1.txtDesc.value			= """ & ConvSPChars(E7_a_allc_rcpt(A295_E5_allc_rcpt_desc))		& """" & vbCr	

	Response.Write " .frm1.txtXchRate.Text			= """ & UNINumClientFormat(E8_f_prrcpt(A295_E4_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtGlNo.value			= """ & ConvSPChars(E5_a_gl)									& """" & vbCr

	Response.Write " .frm1.txtBalAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E8_f_prrcpt(A295_E4_bal_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """" & vbCr
	Response.Write " .frm1.txtBalLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E8_f_prrcpt(A295_E4_loc_bal_amt),gCurrency,ggAmtOfMoneyNo, "X" , "X")	& """" & vbCr
	Response.Write " .frm1.txtClsAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_allc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	& """" & vbCr
	Response.Write " .frm1.txtClsLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_allc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write " .frm1.txtDcAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	& """" & vbCr
	Response.Write " .frm1.txtDcLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """" & vbCr
		
	Response.Write " .frm1.txtDcAmt2.Text			= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	& """" & vbCr
	Response.Write " .frm1.txtDcLocAmt2.Text		= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A295_E5_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """" & vbCr

	Response.Write " End With "						& vbCr
    Response.Write "</Script>"						& vbCr     
    
    
    intCount = UBound(EG1_export_group,1)
    
	If Not IsArray(EG1_export_group) Or IsEmpty(EG1_export_group) Then
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다!
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If
    
	If E4_a_open_ar = EG1_export_group(intCount,A295_EG2_E5_ar_no) Then
		StrNextKey = ""   ' import view
	Else
		StrNextKey = E4_a_open_ar
	End If
	
	If IsEmpty(EG1_export_group) = False And IsArray(EG1_export_group) = True Then    
		lgCurrency = ConvSPChars(EG1_export_group(0,A295_EG2_E5_doc_cur))	
	
		Response.Write "<Script Language=VBScript> " & vbCr
		Response.Write " parent.frm1.hArDocCur.value  = """ & ConvSPChars(EG1_export_group(0,A295_EG2_E5_doc_cur)) & """" & vbCr
		Response.Write "</Script> " & vbCr
		
		For IntRows = 0 To intCount		
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E5_ar_no))
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A295_EG2_E5_ar_dt))    
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows,A295_EG2_E5_ar_due_dt))

			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E5_doc_cur))
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E5_ar_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E5_bal_amt),gCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E3_cls_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E3_cls_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E3_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(IntRows,A295_EG2_E3_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E3_cls_ar_desc))
						
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E4_acct_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E4_acct_nm))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E1_biz_area_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows,A295_EG2_E1_biz_area_nm))

			strData = strData & Chr(11) & LngMaxRow + IntRows
			strData = strData & Chr(11) & Chr(12)  
		Next	
	End If		

    Response.Write "<Script Language=VBScript>  "															& vbCr  
    Response.Write " With parent "																			& vbCr 
    Response.Write " .ggoSpread.Source = .frm1.vspdData1 "													& vbCr
    Response.Write " .ggoSpread.SSShowData """ & strData													& """" & vbCr
    Response.Write " .lgStrPrevKey		 = """ & StrNextKey													& """" & vbCr
    Response.Write " End With "																				& vbCr
    Response.Write "</Script>"  																			& vbCr	    

	If IsArray(EG2_export_group_dc) Or IsEmpty(EG2_export_group_dc) = False Then
		strData	 = ""		
		intCount1 = UBound(EG2_export_group_dc,1)
		For IntRows = 0 To intCount1
    	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A295_EG1_E2_seq))
    	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A295_EG1_E1_acct_cd))
    	    strData = strData & Chr(11) & ""
    	    strData = strData & Chr(11) & ConvSPChars(EG2_export_group_dc(IntRows,A295_EG1_E1_acct_nm))

    	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(IntRows,A295_EG1_E2_dc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
    	    strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(IntRows,A295_EG1_E2_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
    	    	    
            strData = strData & Chr(11) & LngMaxRow + IntRows                                 '11
            strData = strData & Chr(11) & Chr(12)           
		Next
		
	    Response.Write "<Script Language=VBScript> "																& vbCr  
		Response.Write " With parent "																				& vbCr 		
		Response.Write " .ggoSpread.Source = .frm1.vspdData "														& vbCr
		Response.Write " .ggoSpread.SSShowData	""" & strData														& """" & vbCr
		Response.Write " .lgStrPrevKey		=	""" & StrNextKey													& """" & vbCr
		Response.Write " .lgStrPrevKey1		=	""" & StrNextKey1													& """" & vbCr
		Response.Write " End With "																					& vbCr
		Response.Write "</Script>"  																				& vbCr          			
		
		If cint(intCount1) > 0 Then			
			If E3_a_rcpt_dc = EG2_export_group_dc(intCount1,A295_EG1_E2_seq) Then
				StrNextKey1 = ""   ' import view
			Else
				StrNextKey1 = E3_a_rcpt_dc
			End If
			
		End If
	End If
		
    Response.Write "<Script Language=VBScript> "															& vbCr  
    Response.Write "parent.DbQueryOk "																		& vbCr '☜: 모든 쿼리가 종료됨을 알림 
    Response.Write "</Script>"  																			& vbCr          	

%>		
