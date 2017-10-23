<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : REPAY LOAN MULTI DELETE
'*  3. Program ID        : f4255mb3
'*  4. Program 이름      : 차입금멀티상환(삭제)
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
<%																						'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd																		'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next																	'☜: 
Err.Clear 
 
Call LoadBasisGlobalInf()

Dim iPAFG430																			'☆ : 조회용 ComPlus Dll 사용 변수 
Dim strMode																				'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iCommandSent

Dim I1_f_ln_repay
Const A861_I1_repay_no = 0

strMode = Request("txtMode")															'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then												'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
		Call HideStatusWnd		
	ElseIf Trim(Request("txtRePayNo")) = "" Then											'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)						'조회 조건값이 비어있습니다!
		Response.End
		Call HideStatusWnd		 
	End If

	iCommandSent = "DELETE"

	Redim I1_f_ln_repay(6)
	I1_f_ln_repay(A861_I1_repay_no) = Trim(Request("txtRePayNO"))

	Set iPAFG430 = Server.CreateObject("PAFG430.cFMngRepayMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If	

	Call iPAFG430.F_MANAGE_REPAY_MULTI_SVR(gStrGlobalCollection,iCommandSent, I1_f_ln_repay)						
		
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG430 = Nothing															'☜: ComProxy Unload
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If
	
	Set iPAFG430 = Nothing																'☜: Unload Complus	
	                                                
	Response.Write " <Script Language=vbscript> " & vbCr
   	Response.Write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
