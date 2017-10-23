<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : A404MB1
'*  4. Program Name         : PAYMENT 조회하는 P/G
'*  5. Program Desc         : PAYMENT 조회하는 P/G
'*  6. Comproxy List        : +AP004MP
'            
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2000/04/19
'*  9. Modifier (First)     : CHANG SUNG HEE
'* 10. Modifier (Last)      : CHANG SUNG HEE
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/01 : ..........
'**********************************************************************************************
								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


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
On Error Resume Next														'☜: 

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
Dim StrNextKey1
Dim lgStrPrevKey
Dim lgIntFlgMode
Dim test
Dim lgCurrency

Dim I3_a_open_ap 
Dim I2_a_allc_rcpt 
Dim I1_a_rcpt 
Dim E1_b_biz_area 
Dim E2_a_open_ap 
Dim E3_a_rcpt 
Dim E6_a_rcpt 
Dim E7_a_allc_rcpt 
Dim E5_b_biz_partner 
Dim E4_b_acct_dept 
Dim E8_a_gl 
Dim EG1_export_group_assn 
Dim EG2_export_group 
Dim arrCount

Const A297_I2_a_allc_rcpt_allc_no = 0

Const A297_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_area
Const A297_E1_biz_area_nm = 1

Const A297_E4_dept_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_acct_dept
Const A297_E4_dept_nm = 1

Const A297_E5_bp_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_partner
Const A297_E5_bp_nm = 1

Const A297_E7_allc_no = 0    '[CONVERSION INFORMATION]  View Name : export a_allc_rcpt
Const A297_E7_allc_dt = 1
Const A297_E7_allc_type = 2
Const A297_E7_ref_no = 3
Const A297_E7_allc_amt = 4
Const A297_E7_allc_loc_amt = 5
Const A297_E7_dc_amt = 6
Const A297_E7_dc_loc_amt = 7
Const A297_E7_allc_rcpt_desc = 8
Const A297_E7_temp_gl_no = 9

    '[CONVERSION INFORMATION]  Group Name : export_group_assn
Const A297_EG1_E1_dept_cd = 0    '[CONVERSION INFORMATION]  View Name : export_rcpt b_acct_dept
Const A297_EG1_E1_dept_nm = 1
Const A297_EG1_E2_acct_cd = 2    '[CONVERSION INFORMATION]  View Name : export a_acct
Const A297_EG1_E2_acct_nm = 3
Const A297_EG1_E3_allc_dt = 4    '[CONVERSION INFORMATION]  View Name : export a_allc_rcpt_assn
Const A297_EG1_E3_allc_amt = 5
Const A297_EG1_E3_allc_loc_amt = 6
Const A297_EG1_E3_xch_rate = 7
Const A297_EG1_E4_rcpt_no = 8    '[CONVERSION INFORMATION]  View Name : export a_rcpt
Const A297_EG1_E4_rcpt_dt = 9
Const A297_EG1_E4_rcpt_amt = 10
Const A297_EG1_E4_rcpt_loc_amt = 11
Const A297_EG1_E4_allc_amt = 12
Const A297_EG1_E4_allc_loc_amt = 13
Const A297_EG1_E4_bal_amt = 14
Const A297_EG1_E4_bal_loc_amt = 15

    '[CONVERSION INFORMATION]  Group Name : export_group
Const A297_EG2_E1_biz_area_cd = 0    '[CONVERSION INFORMATION]  View Name : export_ap b_biz_area
Const A297_EG2_E1_biz_area_nm = 1
Const A297_EG2_E2_dept_cd = 2    '[CONVERSION INFORMATION]  View Name : export_cls_ap b_acct_dept
Const A297_EG2_E2_dept_nm = 3
Const A297_EG2_E3_bp_cd = 4    '[CONVERSION INFORMATION]  View Name : export_cls_ap_new b_biz_partner
Const A297_EG2_E3_bp_nm = 5
Const A297_EG2_E4_acct_cd = 6    '[CONVERSION INFORMATION]  View Name : export_cls_ap_new a_acct
Const A297_EG2_E4_acct_nm = 7
Const A297_EG2_E5_cls_dt = 8    '[CONVERSION INFORMATION]  View Name : export_new a_cls_ap
Const A297_EG2_E5_doc_cur = 9
Const A297_EG2_E5_xch_rate = 10
Const A297_EG2_E5_cls_amt = 11
Const A297_EG2_E5_cls_loc_amt = 12
Const A297_EG2_E6_cls_ap_desc = 13
Const A297_EG2_E6_ap_no = 14    '[CONVERSION INFORMATION]  View Name : export_new a_open_ap
Const A297_EG2_E6_ap_dt = 15
Const A297_EG2_E6_doc_cur = 16
Const A297_EG2_E6_xch_rate = 17
Const A297_EG2_E6_ap_due_dt = 18
Const A297_EG2_E6_ap_amt = 19
Const A297_EG2_E6_ap_loc_amt = 20
Const A297_EG2_E6_cls_amt = 21
Const A297_EG2_E6_cls_loc_amt = 22
Const A297_EG2_E6_bal_amt = 23
Const A297_EG2_E6_bal_loc_amt = 24

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

   'Redim I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no+4)
   'I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no)   = Trim(Request("txtAllcNo"))
   'I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no+1) = lgAuthBizAreaCd
   'I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no+2) = lgInternalCd
   'I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no+3) = lgSubInternalCd
   'I2_a_allc_rcpt(A297_I2_a_allc_rcpt_allc_no+4) = lgAuthUsrID	

   I2_a_allc_rcpt   = Trim(Request("txtAllcNo"))

	Set pAr0049 = Server.CreateObject("PARG080.cALkUpAllcRcByApSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	If Trim(lgStrPrevKey) <> "" Then
		I3_a_open_ap = lgStrPrevKey
	Else 
		I3_a_open_ap = ""
	End If

	I1_a_rcpt = ""
	'-----------------------
	'Com Action Area
	'-----------------------
	Call pAr0049.A_LOOKUP_ALLC_RCPT_BY_AP_SVR(gStrGlobalCollection, I3_a_open_ap, I2_a_allc_rcpt, I1_a_rcpt, E1_b_biz_area, E2_a_open_ap, E3_a_rcpt, E6_a_rcpt, E7_a_allc_rcpt, E5_b_biz_partner, E4_b_acct_dept, E8_a_gl, EG1_export_group_assn, EG2_export_group) 

	If CheckSYSTEMError(Err,True) = True Then
		Set pAr0049 = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pAr0049 = Nothing

    Response.Write "<Script Language=VBScript> "                                                          & vbCr  
    Response.Write " With parent "                                                                        & vbCr 	

	Response.Write " .frm1.txtRcptNo.Value			= """ & ConvSPChars(EG1_export_group_assn(0,A297_EG1_E4_rcpt_no))			& """" & vbCr
	Response.Write " .frm1.txtRcptDt.text			= """ & UNIDateClientFormat(EG1_export_group_assn(0,A297_EG1_E4_rcpt_dt))	& """" & vbCr
	Response.Write " .frm1.txtAllcDt.text			= """ & UNIDateClientFormat(E7_a_allc_rcpt(A297_E7_allc_dt))				& """" & vbCr
	Response.Write " .frm1.txtBizCd.Value			= """ & ConvSPChars(E1_b_biz_area(A297_E1_biz_area_cd))						& """" & vbCr
	Response.Write " .frm1.txtBizNm.Value		    = """ & ConvSPChars(E1_b_biz_area(A297_E1_biz_area_nm))						& """" & vbCr
	Response.Write " .frm1.txtDeptCd.Value		    = """ & ConvSPChars(E4_b_acct_dept(A297_E4_dept_cd))						& """" & vbCr
	Response.Write " .frm1.txtBpCd.Value			= """ & ConvSPChars(E5_b_biz_partner(A297_E5_bp_cd))						& """" & vbCr
	Response.Write " .frm1.txtBpNm.Value			= """ & ConvSPChars(E5_b_biz_partner(A297_E5_bp_nm))						& """" & vbCr
	Response.Write " .frm1.txtDocCur.value			= """ & ConvSPChars(E6_a_rcpt)												& """" & vbCr
	Response.Write " .frm1.txtXchRate.text			= """ & UNINumClientFormat(EG1_export_group_assn(0,A297_EG1_E3_xch_rate), ggExchRate.DecPoint, 0)	& """" & vbCr
	Response.Write " .frm1.txtGlNo.value			= """ & ConvSPChars(E8_a_gl)												& """" & vbCr
	Response.Write " .frm1.txtTempGlNo.Value		= """ & ConvSPChars(E7_a_allc_rcpt(A297_E7_temp_gl_no))						& """" & vbCr
	Response.Write " .frm1.txtDesc.Value			= """ & ConvSPChars(E7_a_allc_rcpt(A297_E7_allc_rcpt_desc))					& """" & vbCr	

	Response.Write " .frm1.txtBalAmt.text			= """ & UNIConvNumDBToCompanyByCurrency(EG1_export_group_assn(0,A297_EG1_E4_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """" & vbCr
	Response.Write " .frm1.txtBalLocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(EG1_export_group_assn(0,A297_EG1_E4_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write " .frm1.txtClsAmt.text			= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A297_E7_allc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")					& """" & vbCr
	Response.Write " .frm1.txtClsLocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E7_a_allc_rcpt(A297_E7_allc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	
    Response.Write " End With "                                                                           & vbCr
    Response.Write "</Script>"  																		  & vbCr		
	
    intCount = UBound(EG2_export_group,1)
    
	If IsEmpty(EG2_export_group) Or Not IsArray(EG2_export_group) Then    
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다!
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If

    If E2_a_open_ap = EG2_export_group(intCount,A297_EG2_E6_ap_no) Then
		StrNextKey = ""   ' import view
    Else
		StrNextKey = EG2_export_group(intCount,A297_EG2_E6_ap_no)
	End If	
    
	For IntRows = 0 To intCount		
		lgCurrency = ConvSPChars(EG2_export_group(intRows,A297_EG2_E6_doc_cur))
		strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E6_ap_no))
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E4_acct_cd))
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E4_acct_nm))
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E1_biz_area_cd))
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E1_biz_area_nm))
        strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group(intRows,A297_EG2_E6_ap_dt))    
        strData = strData & Chr(11) & UNIDateClientFormat(EG2_export_group(intRows,A297_EG2_E6_ap_due_dt))
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E6_doc_cur))
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group(intRows,A297_EG2_E6_ap_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group(intRows,A297_EG2_E6_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group(intRows,A297_EG2_E5_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
        strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group(intRows,A297_EG2_E5_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
        strData = strData & Chr(11) & ConvSPChars(EG2_export_group(intRows,A297_EG2_E6_cls_ap_desc))        

        strData = strData & Chr(11) & LngMaxRow + IntRows
		strData = strData & Chr(11) & Chr(12)                                    

	Next

	Response.Write "<Script Language=vbscript>"										& vbCr
	Response.Write " With parent "													& vbCr	
	Response.Write " .ggoSpread.Source = .frm1.vspdData "					& vbCr
	Response.Write " .ggoSpread.SSShowData """ & strData				 & """ ,""F""	 " & vbCr
	Response.Write "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1  & ",.C_DOCCUR,.C_APAMT, ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1  & ",.C_DOCCUR,.C_APREMAMT,  ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1  & ",.C_DOCCUR,.C_APCLSAMT,  ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write " .DbQueryOk "													& vbCr
	Response.Write " End With "														& vbCr
	Response.Write "</Script>"

%>
