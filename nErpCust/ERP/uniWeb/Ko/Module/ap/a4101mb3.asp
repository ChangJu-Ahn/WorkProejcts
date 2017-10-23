
<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb3
'*  4. Program Name         : Open Ap 삭제하는 Logic
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
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
<%																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

	On Error Resume Next														'☜: 
	
	Call LoadBasisGlobalInf()

	Dim iPAPG005																'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	Dim ImportData1		

	' -- 권한관리추가 
	Const A114_I11_a_data_auth_data_BizAreaCd = 0
	Const A114_I11_a_data_auth_data_internal_cd = 1
	Const A114_I11_a_data_auth_data_sub_internal_cd = 2
	Const A114_I11_a_data_auth_data_auth_usr_id = 3

	Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

	Redim I11_a_data_auth(3)
	I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	ReDim ImportData1(31)
	ImportData1(0)  = Trim(Request("txtApNo"))

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then										'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
		Call HideStatusWnd			 
	End If

	Set iPAPG005 = Server.CreateObject("PAPG005.cAMngOpenApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	 
	Call iPAPG005.A_MANAGE_OPEN_AP_SVR(gStrGlobalCollection, "DELETE", , ImportData1, , , , I11_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG005 = Nothing		
		Response.End 
	End If
	    
	Set iPAPG005 = Nothing 

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.DbDeleteOk        " & vbcr
	Response.Write "</Script>" & vbcr
%>
