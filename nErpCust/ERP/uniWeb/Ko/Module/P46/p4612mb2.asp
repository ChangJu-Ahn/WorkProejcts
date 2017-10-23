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

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	
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
 	Err.Clear																'��: Protect system from crashing
  	
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
  	
  	If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
		Set oPP4G702 = Nothing		                                                 '��: Unload Comproxy DLL
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
