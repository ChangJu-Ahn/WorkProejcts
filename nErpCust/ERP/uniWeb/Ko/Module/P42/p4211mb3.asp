<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4211mb3
'*  4. Program Name         : 
'*  5. Program Desc         : Insert, Delete, Update Reservation
'*  6. Comproxy List        : PP4C202.cPMngRsvr
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002-07-06	
'*  9. Modifier (First)     : 2002-09-18
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																					'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

Dim oPP4C202				'PP4C202.cPMngRsvr 
Dim iErrorPosition
Dim iBlnMessage
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Err.Clear																		'��: Protect system from crashing

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

Set oPP4C202 = Server.CreateObject("PP4C202.cPMngRsvr")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call oPP4C202.P_MANAGE_RESERVATION(gStrGlobalCollection, _
								   itxtSpread , _
								   iBlnMessage,_
								   iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4C202 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.GetHiddenFocus(" & iErrorPosition & ", 1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPP4C202 IS Nothing) Then
	Set oPP4C202 = Nothing								'��: Unload Comproxy	
End If

If Cbool(iBlnMessage) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.DisplayMsgBox(""17C021"", ""x"", ""x"", ""x"")" & vbCrLF
	Response.Write "</Script>" & vbCrLF
End If
         
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>


