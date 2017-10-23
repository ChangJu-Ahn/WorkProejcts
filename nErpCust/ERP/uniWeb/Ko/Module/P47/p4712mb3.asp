<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4712mb3.asp
'*  4. Program Name			: Save Resource Consumption Result
'*  5. Program Desc			: Confirm Resource Consumption Result (Called By p4712ma1.asp)
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2001/12/06
'*  8. Modified date(Last) 	: 2002/07/18
'*  9. Modifier (First)		: Jeon, Jaehyun
'* 10. Modifier (Last)		: Kang Seong Moon
'* 11. Comment		: Converted by Kang Seong Moon 2002. 7
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

Call HideStatusWnd
Dim oPP4S501
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

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

On Error Resume Next
Err.Clear

set oPP4S501 = Server.CreateObject("PP4S501.cPMngRsrcCnsm")

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4S501 = Nothing	
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF				
	Response.End
End If

call oPP4S501.P_MANAGE_RSC_CONSUMPTION(gStrGlobalCollection, _
										itxtSpread)

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4S501 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

set oPP4S501 = nothing
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "Call Parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF    
Response.End
%>
