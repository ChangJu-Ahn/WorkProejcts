<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID			: p4612mb2.asp
'*  4. Program Name			: Close Order
'*  5. Program Desc			: Close Production Order
'*  6. Dll List				: +PP4G702.cPCnclClsProdOrd
'*  7. Modified date(First) : 2003-08-26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Chen, Jaehyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	
On Error Resume Next														

Dim oPP4G702										'PP4G702.cPCnclClsProdOrd 

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii   

Dim iErrorPosition

'-----------------------------------------------------------
' SQL Server, APS DB Server Information Read
'-----------------------------------------------------------
 	Err.Clear																'☜: Protect system from crashing
  	
  	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count
	             
	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount)
	             
	For ii = 1 To iCUCount
		itxtSpreadArrCount = itxtSpreadArrCount + 1
		itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next

	itxtSpread = Join(itxtSpreadArr,"")
  	
  	Set oPP4G702 = Server.CreateObject("PP4G702.cPCnclClsProdOrd")    
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
	    Response.End 
    End If

         
	Call oPP4G702.P_CANCEL_CLS_PROD_ORDER(gStrGlobalCollection, _
									itxtSpread, _
									iErrorPosition)
  	
  	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Set oPP4G702 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",2)" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
  	
  	Set oPP4G702 = Nothing   
  		
	Response.Write "<Script Language=vbscript>	" & vbCrLF
	Response.Write "	With parent				" & vbCrLF																
	Response.Write "		.DbSaveOk			" & vbCrLF
	Response.Write "	End With				" & vbCrLF
	Response.Write "</Script>					" & vbCrLF
	Response.End
%>
