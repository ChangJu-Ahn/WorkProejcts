<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4111mb3.asp
'*  4. Program Name			: Release Order
'*  5. Program Desc			: Release By Production Order
'*  6. Comproxy List		: PP4G166.cPReleaseProdOrder
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/09/17
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																														'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

Dim oPP4G166
Dim iErrorPosition																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strTxtSpread

Err.Clear																		'☜: Protect system from crashing

strTxtSpread = Request("txtSpread")

Set oPP4G166 = Server.CreateObject("PP4G166.cPReleaseProdOrder")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call oPP4G166.P_RELEASE_BY_ORDER_SVR(gStrGlobalCollection, _
									Request("txtSpread"), _
									 , _
									 , _
									iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4G166 = Nothing
	Response.End
End If

Set oPP4G166 = Nothing								'☜: Unload Comproxy	

        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk(True)" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>