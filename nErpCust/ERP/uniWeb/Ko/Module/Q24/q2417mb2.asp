<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2417MB2
'*  4. Program Name         : Release
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
									
On Error Resume Next
Call HideStatusWnd

Dim strinsp_class_cd
strinsp_class_cd = "S"	''###�׸��� ������ ���Ǻκ�###	

Dim PQIG150																	'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strInspReqNo
Dim strPlantCd
Dim I2_q_inspection_result
	
Dim iCommand
	
Redim I2_q_inspection_result(5)
	
strInspReqNo = UCase(Trim(Request("txtInspReqNo")))
strPlantCd = UCase(Trim(Request("txtPlantCd")))
	
Set PQIG150 = Server.CreateObject("PQIG150.cQMaintReleaseSvr")
	
If CheckSystemError(Err,True) Then						
	Response.End														'��: �����Ͻ� ���� ó���� ������ 
End If
	    
I2_q_inspection_result(0) = 1
I2_q_inspection_result(3) = strinsp_class_cd
I2_q_inspection_result(4) = strPlantCd
I2_q_inspection_result(5) = UNIConvDate(Request("txtReleaseDt"))
	
iCommand = "Confirm"
Call PQIG150.Q_MAINT_INSP_RELEASE_SVR(gstrGlobalCollection,icommand, I2_q_inspection_result, strInspReqNo)

If CheckSystemError(Err,True) Then
	Set PQIG150 = Nothing
	Response.End														'��: �����Ͻ� ���� ó���� ������ 
End If

Set PQIG150 = Nothing                                                  '��: Unload Comproxy
%>
<Script Language=vbscript>
With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
	.DbSaveOk
End With
</Script>