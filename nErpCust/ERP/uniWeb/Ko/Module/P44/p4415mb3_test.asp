<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4415mb3.asp
'*  4. Program Name			: Save Production Results
'*  5. Program Desc			: Confirm Production Results (Called By p4414ma1.asp, p4415ma1.asp)
'*  6. Comproxy List		: +PP4G452.cPCnfmRsltArr
'*  7. Modified date(First)	: 2000/03/30
'*  8. Modified date(Last) 	: 2002/11/26	
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: KHK
'* 11. Comment		:
'**********************************************************************************************

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

Dim oPP4G452												'☆ : 입력/수정용 ComProxy Dll 사용 변수 

Dim strPlantCd											'☆ : Lookup 용 코드 저장 변수 
Dim iErrorPosition										'☆ : Error Position									
Dim iErrorProdtOrdNo, iErrorOprNo, iErrorGoodMvmt		'☆ : Error Return Value
Dim msgStr1, msgStr2

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Const iErrorGoodMvmt_qty = 0
Const iErrorGoodMvmt_trns_item_cd = 1
Const iErrorGoodMvmt_base_unit = 2

	Err.Clear											'☜: Protect system from crashing
	
	strPlantCd = Request("txtPlantCd")
	
	itxtSpread = ""
		             
	iCUCount = Request.Form("txtCUSpread").Count
		             
	itxtSpreadArrCount = -1
		             
	ReDim itxtSpreadArr(iCUCount)
		             
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next
	
	itxtSpread = Join(itxtSpreadArr,"")
	
	Set oPP4G452 = Server.CreateObject("PP4G452_IF.cPCnfmRsltArr")  '2008-03-18 9:06오후 :: hanc

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	Call oPP4G452.P_CONFIRM_RSLT_ARR(gStrGlobalCollection, _
									strPlantCd, _
									"D", _
									itxtSpread, _
									iErrorProdtOrdNo, _
									iErrorOprNo, _
									iErrorPosition, _
									iErrorGoodMvmt)
									   
	Select Case Trim(Cstr(Err.Description))
		
		Case "B_MESSAGE" & Chr(11) & "189614", "B_MESSAGE" & Chr(11) & "189618"	
			If Err.Description = "B_MESSAGE" & Chr(11) & "189614" Then
				Err.Description = "B_MESSAGE" & Chr(11) & "189625"
					
			ElseIf Err.Description = "B_MESSAGE" & Chr(11) & "189618" Then
			 	Err.Description = "B_MESSAGE" & Chr(11) & "189626"	
			End If
			msgStr1 = "오더번호 : " & iErrorProdtOrdNo & " " & _
					  "공정 : " & iErrorOprNo & VbCrLf
			msgStr2 = "부품 : " & iErrorGoodMvmt(iErrorGoodMvmt_trns_item_cd) & "  " & _
					   UniNumClientFormat(iErrorGoodMvmt(iErrorGoodMvmt_qty),ggQty.DecPoint,0) & " " & iErrorGoodMvmt(iErrorGoodMvmt_base_unit)		   
					   
			If CheckSYSTEMError2(Err,True,msgStr1,msgStr2,"","","") = True  Then
				Set oPP4G452 = Nothing
				If iErrorPosition <> 0 Then
					Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
					Response.Write "Call parent.GetHiddenFocus(" & iErrorPosition & ", 1)" & vbCrLF
					Response.Write "</Script>" & vbCrLF
				End If
				Response.End
			End If
		Case Else
			If CheckSYSTEMError(Err,True) = True Then
				Set oPP4G452 = Nothing
				If iErrorPosition <> 0 Then
					Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
					Response.Write "Call parent.GetHiddenFocus(" & iErrorPosition & ", 1)" & vbCrLF
					Response.Write "</Script>" & vbCrLF
				End If
				Response.End
			End If
	End Select
	If Not(oPP4G452 Is Nothing) Then 
		Set oPP4G452 = Nothing
	End If	

	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
%>






