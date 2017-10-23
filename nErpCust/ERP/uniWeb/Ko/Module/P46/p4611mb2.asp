<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4611mb2.asp
'*  4. Program Name         : List Production Order (Query)
'*  5. Program Desc         : List ASP used by Order Closing
'*  6. Comproxy List        : DB Agent (p4611mb1)
'*  7. Modified date(First) : 2000/03/28
'*  8. Modified date(Last)  : 2003/03/23
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              : 
'**********************************************************************************************%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd

Dim oPP4G701											'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iErrorPosition
Dim iErrorProdtOrdHdr
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount
Dim msgStr1, msgStr2

Const iErrorProdt_Order_No = 0
Const iErrorLatest_Rcpt_Date = 1

Dim iCUCount

Dim ii       

On Error Resume Next

Err.Clear
	
	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count
	             
	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount)
	             
	For ii = 1 To iCUCount
		itxtSpreadArrCount = itxtSpreadArrCount + 1
		itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next

	itxtSpread = Join(itxtSpreadArr,"")
	
	Set oPP4G701 = Server.CreateObject("PP4G701.cPCloseProdOrd")    
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
	    Response.End 
    End If

	Call oPP4G701.P_CLOSE_PROD_ORDER(gStrGlobalCollection, _
									itxtSpread, _
									iErrorProdtOrdHdr, _
									iErrorPosition)

	Select Case Trim(Cstr(Err.Description))
		
		Case "B_MESSAGE" & Chr(11) & "189228"
			msgStr1 = "오더번호 : " & iErrorProdtOrdHdr(iErrorProdt_Order_No)
			msgStr2 = "입고일자 : " & UNIDateClientFormat(iErrorProdtOrdHdr(iErrorLatest_Rcpt_Date))
			If CheckSYSTEMError2(Err,True,msgStr1,msgStr2,"","","") = True  Then
				Set oPP4G701 = Nothing		                                                 '☜: Unload Comproxy DLL
				If iErrorPosition <> 0 Then
					Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
					Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
					Response.Write "</Script>" & vbCrLF
				End If
				Response.End
			End If
		Case Else
			If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
				Set oPP4G701 = Nothing		                                                 '☜: Unload Comproxy DLL
				Response.Write "<Script Language=VBScript>" & vbCrLF
				Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
				Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
				Response.Write "</Script>" & vbCrLF
				Response.End
			End If
	End Select	

	Set oPP4G701 = Nothing                                                   '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCrLF
	Response.Write "	With parent				" & vbCrLF																
	Response.Write "		.DbSaveOk			" & vbCrLF
	Response.Write "	End With				" & vbCrLF
	Response.Write "</Script>					" & vbCrLF
	Response.End
%>	