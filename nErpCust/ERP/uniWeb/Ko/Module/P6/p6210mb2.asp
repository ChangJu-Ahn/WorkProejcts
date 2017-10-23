<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID           : p6210mb2.asp
'*  4. Program Name         : Save Cast Result
'*  5. Program Desc			: 
'*  6. Comproxy List		: 
'*  7. Modified date(First) : 2005/10/18
'*  8. Modified date(Last)  : 2005/10/18
'*  9. Modifier (First)     : Chen, Jae Hyun
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd
On Error Resume Next
Dim oPY6G230									'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strPlantCd
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear										'☜: Protect system from crashing

strPlantCd = UCase(Request("txtPlantCd"))

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Set oPY6G230 = Server.CreateObject("PY6G230.cPMngCastRslt")
If CheckSYSTEMError(Err,True) = True Then
	Set oPY6G230 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF					
	Response.End
End If

call oPY6G230.P_MANAGE_CAST_RSLT(gStrGlobalCollection, _
								strPlantCd, _
								itxtSpread, _
								iErrorPosition)
								
If  CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
	Set oPY6G230 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Set oPY6G230 = Nothing															'☜: Unload Comproxy

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
