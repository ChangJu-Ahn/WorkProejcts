
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5407mb1
'*  4. Program Name         : 미결반제(만기어음)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2002/11/05
'*  8. Modified date(Last)  : 2002/11/05
'*  9. Modifier (First)     : KIM HO YOUNG
'* 10. Modifier (Last)      : KIM HO YOUNG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2002/11/05 : ..........
'**********************************************************************************************


Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
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
'	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
' 수정을 요함 
Dim pGlCardAcct																'☆ : 조회용 ComProxy Dll 사용 변수 

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
Dim I1_mgnt_acct_cd
Dim I2_b_bank
Dim I3_trans_type
Dim I4_a_gl
Dim I5_gCurrency
Dim arrCount
Dim lgCurrency
Dim E1_cls_no
' Com+ Conv. 변수 선언 
Dim pvStrGlobalCollection 

Const I2_b_bank_cd = 0
Const I2_b_bank_acct = 1

Const I4_gl_dt = 0
Const I4_gl_desc = 1
Const I4_gl_input_type = 2
Const I4_dept_cd = 3
Const I4_org_change_id = 4

Dim I6_txtFromBaseDt
Dim I7_txtToBaseDt

ReDim I2_b_bank(I2_b_bank_acct)    
ReDim I4_a_gl(I4_org_change_id)

'[CONVERSION INFORMATION]  View Name : export a_open_acct



'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################
lgStrPrevKey = Request("lgStrPrevKey")

I1_mgnt_acct_cd = Trim(Request("hmgnt_acct_cd")) ' 당좌예금의 계정코드 

I2_b_bank(I2_b_bank_cd) = Trim(Request("htxtBankCd"))
I2_b_bank(I2_b_bank_acct) = Trim(Request("htxtBankAcct"))

I3_trans_type = "AP011"

I4_a_gl(I4_gl_dt) = UNIConvDate(Request("htxtGlDt"))
I4_a_gl(I4_gl_desc) = Trim(Request("htxtDesc"))
I4_a_gl(I4_gl_input_type) = "OC"
I4_a_gl(I4_dept_cd) = Trim(Request("htxtDeptCd"))
I4_a_gl(I4_org_change_id) = Trim(Request("htxtOrgChangeId"))

I5_gCurrency = gCurrency


'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################


Set pGlCardAcct = Server.CreateObject("PAUG035.cACreGlCardAcctSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Set pOpenCardAcct = Nothing												'☜: ComProxy Unload
	Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'⊙:
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

	LngMaxRow  = CLng(Request("txtMaxRows"))												'☜: Fetechd Count      
	LngMaxRow1  = CLng(Request("txtMaxRows1"))

	I6_txtFromBaseDt = UniConvDateAToB(Trim(Request("txtFromBaseDt")),gDateFormat, gServerDateFormat)
	I7_txtToBaseDt = UniConvDateAToB(Trim(Request("txtToBaseDt")),gDateFormat, gServerDateFormat)


	Call pGlCardAcct.A_CREATE_GL_CARD_ACCT_SVR(gStrGlobalCollection,I1_mgnt_acct_cd,I2_b_bank,I3_trans_type,_
						I4_a_gl,I5_gCurrency,I6_txtFromBaseDt, I7_txtToBaseDt, E1_cls_no)
'-----------------------
'Com Action Area
'-----------------------

	If CheckSYSTEMError(Err,True) = True Then
	
		Set pGlCardAcct = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pGlCardAcct = Nothing
    
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
	Response.Write " .frm1.txtClsNo.value = """ & ConvSPChars(Trim(E1_cls_no)) & """" & vbCr
	Response.Write " .DbSaveOK() "								  & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"     	
%>		

