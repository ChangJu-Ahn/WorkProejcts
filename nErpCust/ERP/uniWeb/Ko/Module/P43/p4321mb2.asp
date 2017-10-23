<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4321mb2.asp
'*  4. Program Name			: Save BackLog
'*  5. Program Desc			: Confirm BackLog
'*  6. Comproxy List		: +PBATP463
'*  7. Modified date(First)	: 2006-04-11
'*  8. Modified date(Last) 	:
'*  9. Modifier (First)		:HJO
'* 10. Modifier (Last)		: 
'* 11. Comment		:
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")

Call HideStatusWnd

On Error Resume Next

Dim oPI0C290

Dim iErrorPosition										'☆ : Error Position									

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

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

Set oPI0C290 = Server.CreateObject("PBATP463.cPBacklogMain")
	
'----------Developer Coding part (Start)--------------------------------------------------------------
	Call oPI0C290.P_ISSUE_BACKLOG_MAIN(gStrGlobalCollection, itxtSpread)
									   
'	Select Case Trim(Cstr(Err.Description))		
'----------Developer Coding part (End)--------------------------------------------------------------				

	If CheckSYSTEMError(Err,True) = True Then
		Set oPI0C290 = Nothing
		If iErrorPosition <> 0 Then
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
			Response.Write "</Script>" & vbCrLF
		End If
		Response.End
	End If
	'End Select

	Set oPI0C290 = Nothing
	
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End	
	%>
