<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: P4214mb2.asp
'*  4. Program Name			: ManageOrdDoc
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP4G207.cPMngOrdDoc
'*  7. Modified date(First)	: 
'*  8. Modified date(Last) 	: 2002-07-02
'*  9. Modifier (First)		: 
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																					'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next																'��: 

Dim pPP4G207																			'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strCode																			'�� : Lookup �� �ڵ� ���� ���� 
Dim strProdOrderNo																	'�� : Lookup �� �ڵ� ���� ����									
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Err.Clear																		'��: Protect system from crashing

strMode = Request("txtMode")														'�� : ���� ���¸� ���� 

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

Set pPP4G207 = Server.CreateObject("PP4G207.cPMngOrdDoc")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
Call pPP4G207.P_MANAGE_ORD_DOC(gStrGlobalCollection, _
							   itxtSpread, _
							   iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set pPP4G207 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF	
	Response.End
End If

Set pPP4G207 = Nothing	

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>