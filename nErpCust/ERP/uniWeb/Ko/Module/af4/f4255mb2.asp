<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : REPAY LOAN MULTI SAVE
'*  3. Program ID        : f4255mb2
'*  4. Program 이름      : 차입금멀티상환(저장)
'*  5. Program 설명      : 차입금멀티상환 
'*  6. Complus 리스트    : PAFG430.DLL
'*  7. 최초 작성년월일   : 2003/05/10
'*  8. 최종 수정년월일   : 2003/05/10
'*  9. 최초 작성자       : 정용균 
'* 10. 최종 작성자       : 정용균 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																								'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd()																			'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next																			'☜: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim iPAFG430																					'저장용 ComPlus Dll 사용 변수 
Dim lgIntFlgMode
Dim iCommandSent 

Dim I1_f_ln_repay
Const A861_I1_repay_no = 0
Const A861_I1_repay_dt = 1
Const A861_I1_repay_dept_cd = 2
Const A861_I1_repay_org_change_id = 3
Const A861_I1_repay_user_fld1 = 4
Const A861_I1_repay_user_fld2 = 5
Const A861_I1_repay_desc = 6

Dim E1_b_auto_numbering 

	lgIntFlgMode = CInt(Request("txtMode"))														'☜: 저장시 Create/Update 판별 

	'-----------------------
	'Data manipulate area
	'---------- -------------																	'Single 데이타 저장 
	ReDim I1_f_ln_repay(A861_I1_repay_desc)
	
	I1_f_ln_repay(A861_I1_repay_no) = Trim(Request("txtRePayNO"))
	I1_f_ln_repay(A861_I1_repay_dt) = UNIConvDate(Request("txtRePayDT"))
	I1_f_ln_repay(A861_I1_repay_org_change_id) = UCase(Request("horgChangeId"))
	I1_f_ln_repay(A861_I1_repay_dept_cd) = UCase(Trim(Request("txtDeptCd")))
	I1_f_ln_repay(A861_I1_repay_user_fld1) = Trim(Request("txtUserFld1"))
	I1_f_ln_repay(A861_I1_repay_user_fld2) = Trim(Request("txtUserFld2"))
	I1_f_ln_repay(A861_I1_repay_desc) = Trim(Request("txtRePayDesc"))

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set iPAFG430 = Server.CreateObject("PAFG430.cFMngRepayMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		

	Call iPAFG430.F_MANAGE_REPAY_MULTI_SVR(gStrGlobalCollection,iCommandSent, I1_f_ln_repay, _
								Trim(Request("txtSpread4")),Trim(Request("txtSpread1")), _
								Trim(Request("txtSpread")), E1_b_auto_numbering)						

	'---------------------------------------------
	'Com action result check area(OS,internal)
	'---------------------------------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG430 = Nothing																	'☜: ComPlus Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		
		
	Set iPAFG430 = Nothing																		'☜: ComPlus Unload		


    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_b_auto_numbering)  & """)" & vbCr
    Response.Write "End With "					 & vbCr	  
    Response.Write "</Script>"           
%>
