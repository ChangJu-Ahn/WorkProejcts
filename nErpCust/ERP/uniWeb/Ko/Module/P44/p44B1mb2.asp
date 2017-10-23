<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p44B1mb2.asp
'*  4. Program Name			: Save POP Result
'*  5. Program Desc			: Release By Production Order
'*  6. Comproxy List		: PBATP442.cUpdPopInf
'*  7. Modified date(First) : 2005/12/15
'*  8. Modified date(Last)  : 2005/12/15
'*  9. Modifier (First)     : Chen, Jae Hyun
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

On Error Resume Next

Dim oPBATP442
Dim iErrorPosition
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear																		'��: Protect system from crashing

iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Set oPBATP442 = Server.CreateObject("PBATP442.cUpdPopInf")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
Call oPBATP442.UPD_POP_MAIN(gStrGlobalCollection, _
						itxtSpread)

If CheckSYSTEMError(Err, True) = True Then
	Set oPBATP442 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "parent.DbSaveNotOk" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPBATP442 Is nothing) Then
	Set oPBATP442 = Nothing								'��: Unload Comproxy	
End If	

        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
