<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4512mb2.asp
'*  4. Program Name			: Cancel Goods Receipt For Production Order
'*  5. Program Desc			: 생산입고취소 
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

Dim oPP4G653
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iDCount

Dim ii

itxtSpread = ""

iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iDCount)
             
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Err.Clear																		'☜: Protect system from crashing

 	Set oPP4G653 = CreateObject("PP4G653.cPCnclProdGoodsRcpt")    
    If CheckSYSTEMError(Err,True) = True Then
		Set oPP4G653 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF					
		Response.End
	End If
	
	Call oPP4G653.P_CANCEL_PROD_GOODS_RECEIPT(gStrGlobalCollection, _
												itxtSpread, _
												iErrorPosition)
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Set oPP4G653 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF						
		Response.End
	End If
	
	Set oPP4G653 = nothing
    
    Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF	
	Response.End
%>
