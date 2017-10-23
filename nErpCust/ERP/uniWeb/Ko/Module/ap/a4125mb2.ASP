<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : Multi Alloction Payment CUD
'*  3. Program ID        : a4125mb2
'*  4. Program 이름      : 멀티출금(저장)
'*  5. Program 설명      : 멀티출금 저장(신규,수정)
'*  6. Complus 리스트    : PAPG070
'*  7. 최초 작성년월일   : 2003/03/27
'*  8. 최종 수정년월일   : 2003/03/27
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
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd()														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim iPAPG070																'저장용 ComPlus Dll 사용 변수 
Dim lgIntFlgMode

Dim iCommandSent 

Dim I1_a_allc_Paym
Const A389_I1_allc_no = 0
Const A389_I1_allc_dt = 1
Const A389_I1_org_change_id = 2
Const A389_I1_dept_cd = 3
Const A389_I1_bp_cd = 4
Const A389_I1_allc_type = 5
Const A389_I1_allc_amt = 6
Const A389_I1_allc_loc_amt = 7
Const A389_I1_dc_amt = 8
Const A389_I1_dc_loc_amt = 9
Const A389_I1_allc_Paym_desc = 10

Dim I2_b_currency 

Dim txtSpread_allc 
Dim txtSpread_cls 
Dim txtSpread_etc 
Dim txtSpread_etc_dtl 

Dim E1_auto_no 

	lgIntFlgMode = CInt(Request("txtMode"))									'☜: 저장시 Create/Update 판별 

	'-----------------------
	'Data manipulate area
	'---------- -------------												    'Single 데이타 저장 
	ReDim I1_a_allc_Paym(A389_I1_allc_Paym_desc)
	
	I1_a_allc_Paym(A389_I1_allc_no) = Trim(Request("txtPaymNo"))
	I1_a_allc_Paym(A389_I1_allc_dt) = UNIConvDate(Request("txtPaymDt"))
	I1_a_allc_Paym(A389_I1_org_change_id) = UCase(Request("hOrgChangeId"))
	I1_a_allc_Paym(A389_I1_dept_cd) = UCase(Trim(Request("txtDeptCd")))
	I1_a_allc_Paym(A389_I1_bp_cd) = UCase(Trim(Request("txtBpCd")))
	I1_a_allc_Paym(A389_I1_allc_type) = "X"
	I1_a_allc_Paym(A389_I1_allc_amt) = 0
	I1_a_allc_Paym(A389_I1_allc_loc_amt) = 0
	I1_a_allc_Paym(A389_I1_dc_amt) = 0
	I1_a_allc_Paym(A389_I1_dc_loc_amt) = 0
	I1_a_allc_Paym(A389_I1_allc_Paym_desc) = ""'Trim(Request("txtPaymDesc"))

	I2_b_currency = gCurrency

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	Set iPAPG070 = Server.CreateObject("PAPG070.cAMntMultiPaymSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		

	Call iPAPG070.A_MAINT_MULTI_PAYM_SVR(gStrGlobalCollection,iCommandSent, I1_a_allc_Paym, _
								I2_b_currency,Trim(Request("txtSpread4")),Trim(Request("txtSpread1")), _
								Trim(Request("txtSpread")),Trim(Request("txtSpread3")), E1_auto_no)						
																		'☜: 비지니스 로직 처리를 종료함 
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG070 = Nothing																	'☜: ComPlus Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		
		
	Set iPAPG070 = Nothing																		'☜: ComPlus Unload		


    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_auto_no)  & """)" & vbCr
    Response.Write "End With "					 & vbCr	  
    Response.Write "</Script>"           
%>
