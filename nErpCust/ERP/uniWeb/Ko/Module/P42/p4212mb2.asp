<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4212mb2.asp
'*  4. Program Name			: Release Order
'*  5. Program Desc			: Release By Production Order
'*  6. Comproxy List		: PP4G166.cPRlseByOrdSvr
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2003/02/04
'*  9. Modifier (First)     : Park, BumSoo
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

Dim oPP4G166
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

Set oPP4G166 = Server.CreateObject("PP4G166.cPReleaseProdOrder")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
Call oPP4G166.P_RELEASE_BY_ORDER_SVR(gStrGlobalCollection, _
									 itxtSpread, _
									 , _
									 , _
									iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set oPP4G166 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",2)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPP4G166 Is nothing) Then
	Set oPP4G166 = Nothing								'��: Unload Comproxy	
End If	

        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
