<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: B1B06MB2.asp
'*  4. Program Name			: Basic Data
'*  5. Program Desc			: Save B_ITEM_ACCT
'*  6. Comproxy List		: PB3S115.cBSetItemAcct
'*  7. Modified date(First)	: 2004/12/01
'*  8. Modified date(Last) 	: 2004/12/01
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%														'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

Call HideStatusWnd

On Error Resume Next

Dim oPB3S115
Dim iErrorPosition																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim itxtSpread

Err.Clear																		'☜: Protect system from crashing

itxtSpread = ""

itxtSpread = Request("txtSpread")

Set oPB3S115 = Server.CreateObject("PB3S115.cBSetItemAcct")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call oPB3S115.B_MANAGE_ITEM_ACCT(gStrGlobalCollection, _
								itxtSpread, _
								iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPB3S115 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPB3S115 Is nothing) Then
	Set oPB3S115 = Nothing								'☜: Unload Comproxy	
End If
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "Call parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>