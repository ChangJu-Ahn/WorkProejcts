<%
'**********************************************************************************************
'*  1. Module Name			: Interface 
'*  2. Function Name		: 
'*  3. Program ID			: xi313mb3_ko119.asp
'*  4. Program Name			: �����������
'*  5. Program Desc			: 
'*  6. Comproxy List		: +PXI3G132_KO119
'*  7. Modified date(First)	: 2006-04-24
'*  8. Modified date(Last) 	:
'*  9. Modifier (First)		:HJO
'* 10. Modifier (Last)		: 
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

Call HideStatusWnd

On Error Resume Next

    Session.Timeout = 60          ' minute 
    Server.ScriptTimeOut = 3600   ' NumSeconds

Dim strPlantCd											'�� : Lookup �� �ڵ� ���� ���� 
Dim iErrorPosition										'�� : Error Position									
Dim iErrorProdtOrdNo, iErrorOprNo, iErrorGoodMvmt		'�� : Error Return Value
Dim msgStr1, msgStr2

Dim oPXI3122

Dim iCUCount

Dim ii											'�� : Lookup �� �ڵ� ���� ���� 

	Err.Clear											'��: Protect system from crashing
	
	Set oPXI3122 = Server.CreateObject("PXI3G132_KO119.cUploadErpProdRcpt")
	If CheckSYSTEMError(Err,True) = True Then
		Set oPXI3122 = Nothing
		Response.End
	End If		

	Call oPXI3122.UPLOAD_ERP_PROD_RCPT_MAIN(gStrGlobalCollection,"")

	If CheckSYSTEMError(Err,True) = True Then
		Set oPXI3122 = Nothing
		Response.End
	
	End If
	Set oPXI3122 = Nothing
		
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End	
	%>
