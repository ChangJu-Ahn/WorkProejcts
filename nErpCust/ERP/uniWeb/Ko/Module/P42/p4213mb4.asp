<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4213mb4.asp
'*  4. Program Name			: Cancel Release Production Order
'*  5. Program Desc			: Cancel Release Production Order
'*  6. Comproxy List		: PP4G255.cPCnclRlse
'*  7. Modified date(First) : 2000/05/25
'*  8. Modified date(Last)  : 2003/02/04
'*  9. Modifier (First)     : Im, Hyun Soo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
																	
Call HideStatusWnd

On Error Resume Next

Dim oPP4G255						'PP4G255.cPCnclRlse
Dim iErrorPosition
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear																		'��: Protect system from crashing

Set oPP4G255 = Server.CreateObject("PP4G255.cPCnclRlse")

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)
             
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
		
Call oPP4G255.P_CANCEL_RELEASE(gStrGlobalCollection, _
							   itxtSpread, _
							   iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set oPP4G255 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",2)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPP4G255 Is nothing) Then
	Set oPP4G255 = Nothing								'��: Unload Comproxy	
End If
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>

