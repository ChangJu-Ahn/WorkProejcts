
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb3
'*  4. Program Name         : 기초 Open Ap 삭제하는 Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1															'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

	On Error Resume Next														'☜: 

	Call LoadBasisGlobalInf()

	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

	Dim iPARG005																'조회용 ComPlus Dll 사용 변수 
	Dim iArrData
	Dim ImportTypeTransType

	Const OpenArNo = 0

	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then										'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End
		Call HideStatusWnd		
	End If

	ImportTypeTransType = "AR005"

	Redim iarrdata(28)    
	iArrData(OpenArNo)  = Trim(Request("txtArNo"))

	Set iPARG005 = Server.CreateObject("PARG005.cAMngOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	   
	Call iPARG005.A_MANAGE_OPEN_AR_SVR (gStrGlobalCollection, "DELETE", ImportTypeTransType, _ 
	                                           , , , , , , , , iArrData)
	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG005 = Nothing		
		Response.End 
	End If
	    
	Set iPARG005 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbDeleteOk        " & vbcr
	Response.Write "</Script>" & vbcr
%>

