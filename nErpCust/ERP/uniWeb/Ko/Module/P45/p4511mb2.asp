<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4511mb2.asp
'*  4. Program Name			: Goods Receipt For Production Order
'*  5. Program Desc			: Goods Receipt For Production Order
'*  6. Comproxy List		: 
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2002/07/18
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Kang Seong Moon
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd
On Error Resume Next
Dim oPP4G651									'☆ : 입력/수정용 ComProxy Dll 사용 변수 
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

Set oPP4G651 = Server.CreateObject("PP4G651.cPProdGoodsRcpt")
If CheckSYSTEMError(Err,True) = True Then
	Set oPP4G651 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF					
	Response.End
End If

call oPP4G651.P_PROD_GOODS_RCPT(gStrGlobalCollection, _
								strPlantCd, _
								itxtSpread, _
								iErrorPosition)
								
If  CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
	Set oPP4G651 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Set oPP4G651 = Nothing															'☜: Unload Comproxy

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
