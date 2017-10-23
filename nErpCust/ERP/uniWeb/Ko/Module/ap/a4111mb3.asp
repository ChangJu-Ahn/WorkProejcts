<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4111mb3
'*  4. Program Name         : 채무/채권 상계 삭제 Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/10
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
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
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim pAp0081																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iCommandSent
Dim I1_a_acct_trans_type
Dim I2_b_acct_dept
Dim I4_a_clear_ap_ar

Const A360_I2_org_change_id = 0    
Const A360_I2_dept_cd = 1

Const A360_I4_clear_no = 0    
Const A360_I4_clear_dt = 1
Const A360_I4_ref_no = 2
Const A360_I4_clear_amt = 3
Const A360_I4_clear_loc_amt = 4
Const A360_I4_internal_cd = 5
Const A360_I4_insrt_user_id = 6
Const A360_I4_insrt_dt = 7
Const A360_I4_updt_user_id = 8
Const A360_I4_updt_dt = 9
Const A360_I4_doc_cur = 10

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

	Dim I5_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A360_I5_a_data_auth_data_BizAreaCd = 0
	Const A360_I5_a_data_auth_data_internal_cd = 1
	Const A360_I5_a_data_auth_data_sub_internal_cd = 2
	Const A360_I5_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I5_a_data_auth(3)
	I5_a_data_auth(A360_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	End If

	ReDim I2_b_acct_dept(A360_I2_dept_cd)
	ReDim I4_a_clear_ap_ar(A360_I4_doc_cur)

	iCommandSent = "DELETE"
	I4_a_clear_ap_ar(A360_I4_clear_no) = Trim(Request("txtClearNo"))

	Set pAp0081 = Server.CreateObject("PAPG055.cAMntClearApArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If	

	Call pAp0081.A_MAINT_CLEAR_AP_AR_SVR(gStrGlobalCollection,iCommandSent,,,,I4_a_clear_ap_ar,,,I5_a_data_auth)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAp0081 = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If	

	Set pAp0081 = Nothing                                                   '☜: Unload Comproxy

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr

%>
