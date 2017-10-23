<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4311mb3.asp
'*  4. Program Name			: Issue Components Manually
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP4G301.cPIssCmpsManually
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/07/04
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd


On Error Resume Next

Dim pPP4G303
Dim iErrorPosition
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear																			'☜: Protect system from crashing
	'-----------------------
	'Com Action Area
	'-----------------------
	itxtSpread = ""
             
	iCUCount = Request.Form("txtCUSpread").Count
	             
	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount)
	
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next

	itxtSpread = Join(itxtSpreadArr,"")
	
	Set pPP4G303 = Server.CreateObject("PP4G303_KO441.cPIssCmpsManually")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	Call pPP4G303.P_ISSUE_COMPONENTS_MANUALLY(gStrGlobalCollection, _
											  itxtSpread, _
											  iErrorPosition)
											  
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP4G303 = Nothing															'☜: Unload Component
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "Call parent.GetHiddenFocus(" & iErrorPosition & ", parent.C_IssueQty)" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	Set pPP4G303 = Nothing      			
	
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
%>
