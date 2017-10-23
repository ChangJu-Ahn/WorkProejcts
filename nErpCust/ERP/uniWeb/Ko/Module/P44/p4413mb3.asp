<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4413mb3.asp
'*  4. Program Name			: Save Production Results
'*  5. Program Desc			: Confirm Production Results By Order
'*  6. Comproxy List		: +PP4G452.cPCnfmRsltArr
'*  7. Modified date(First)	: 2000/03/30
'*  8. Modified date(Last) 	: 2002/10/07
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
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

														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

On Error Resume Next

Dim oPP4G452												'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strCode												'�� : Lookup �� �ڵ� ���� ���� 
Dim strPlantCd											'�� : Lookup �� �ڵ� ���� ���� 
Dim iErrorPosition										'�� : Error Position									
Dim iErrorProdtOrdNo, iErrorOprNo, iErrorGoodMvmt		'�� : Error Return Value
Dim msgStr1, msgStr2
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Const iErrorGoodMvmt_qty = 0
Const iErrorGoodMvmt_trns_item_cd = 1
Const iErrorGoodMvmt_base_unit = 2

    Err.Clear											'��: Protect system from crashing

    strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
    
	strPlantCd = Request("txtPlantCd")

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

    Set oPP4G452 = CreateObject("PP4G452.cPCnfmRsltArr")
    
    If CheckSYSTEMError(Err,True) = True Then
	
		Set oPP4G452 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	'Third value is 
	'Case Result By Order: H
	'Case Result By Opr: D
	
	Call oPP4G452.P_CONFIRM_RSLT_ARR(gStrGlobalCollection, _
									 strPlantCd, _
									 "H", _
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
			msgStr1 = "������ȣ : " & iErrorProdtOrdNo & " " & _
					  "���� : " & iErrorOprNo & VbCrLf
			msgStr2 = "��ǰ : " & iErrorGoodMvmt(iErrorGoodMvmt_trns_item_cd) & "  " & _
					   UniNumClientFormat(iErrorGoodMvmt(iErrorGoodMvmt_qty),ggQty.DecPoint,0) & " " & iErrorGoodMvmt(iErrorGoodMvmt_base_unit)		   
					   
			If CheckSYSTEMError(Err,True) = True Then
				Set oPP4G452 = Nothing
				If iErrorPosition <> 0 Then
					Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
					Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
					Response.Write "</Script>" & vbCrLF
				End If
				Response.End
			End If
		Case Else
			If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
				Set oPP4G452 = Nothing
				If iErrorPosition <> 0 Then
					Response.Write "<Script Language=VBScript>" & vbCrLF
					Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
					Response.Write "Call parent.SheetFocus(" & iErrorPosition & ", 1)" & vbCrLF
					Response.Write "</Script>" & vbCrLF
				End If
				Response.End
			End If
	End Select	
	
	If not(oPP4G452 is nothing)  Then
		Set oPP4G452 = Nothing
	End If
	
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "parent.DbSaveOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End	
%>
