<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: P4214mb2.asp
'*  4. Program Name			: ManageOrdDoc
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP4G207.cPMngOrdDoc
'*  7. Modified date(First)	: 
'*  8. Modified date(Last) 	: 2002-07-02
'*  9. Modifier (First)		: 
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																					'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next																'☜: 

Dim pPP4G207																			'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strCode																			'☆ : Lookup 용 코드 저장 변수 
Dim strProdOrderNo																	'☆ : Lookup 용 코드 저장 변수									
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Err.Clear																		'☜: Protect system from crashing

strMode = Request("txtMode")														'☜ : 현재 상태를 받음 

Err.Clear																		'☜: Protect system from crashing

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount + iDCount)
             
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Set pPP4G207 = Server.CreateObject("PP4G207.cPMngOrdDoc")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
Call pPP4G207.P_MANAGE_ORD_DOC(gStrGlobalCollection, _
							   itxtSpread, _
							   iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPP4G207 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF	
	Response.End
End If

Set pPP4G207 = Nothing	

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>