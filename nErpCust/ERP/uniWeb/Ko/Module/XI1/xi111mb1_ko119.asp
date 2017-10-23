
<%
'**********************************************************************************************
'*  1. Module Name          : interface
'*  2. Function Name        : 
'*  3. Program ID           : xi111mb1_ko119.asp
'*  4. Program Name         :INTERFACE SET (Query)
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/04/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter ���� 
Dim strQryMode

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey										' ���� �� 
Dim lgStrPrevKey									' ���� �� 
Dim LngMaxRow										' ���� �׸����� �ִ�Row
Dim LngRow1
Dim GroupCount1
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strFlag
Dim LngRow

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

	'lgStrPrevKey = Request("lgStrPrevKey")
    
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=====================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "XI111MB1S_KO119"	'main query change id	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtSystemId1")),"''","S")

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	Response.Write "<Script Language=VBScript>" & vbCRLF
	Response.Write "With parent.frm1" & vbCRLF
	Response.Write ".txtSystemId1.value = """ & ConvSPChars(rs0("system_id")) & """" & vbCRLF
	Response.Write ".txtSystemIdNm1.value = """ & ConvSPChars(rs0("system_nm")) & """" & vbCRLF
	Response.Write ".txtSystemId2.value = """ & ConvSPChars(rs0("system_id")) & """" & vbCRLF
	Response.Write ".txtSystemIdNm2.value = """ & ConvSPChars(rs0("system_nm")) & """" & vbCRLF
	Response.Write ".txtPlantCd.value = """ & ConvSPChars(rs0("plant_cd")) & """" & vbCRLF
	Response.Write ".txtPlantNm.value = """ & ConvSPChars(rs0("plant_nm")) & """" & vbCRLF
	If Trim(rs0("usage_flag")) = "Y" Then
		Response.Write ".rdoFlg1.checked = True" & vbCRLF
		Response.Write "parent.lgRdoOldVal1 = 1" & vbCRLF
	Else
		Response.Write ".rdoFlg2.checked = True" & vbCRLF
		Response.Write "parent.lgRdoOldVal1 = 2" & vbCRLF
	End If
	Response.Write ".txtAliasNm.value = """ & ConvSPChars(rs0("alias_nm")) & """" & vbCRLF
	Response.Write ".txtIPAdd.value = """ & ConvSPChars(rs0("ip_address")) & """" & vbCRLF
	Response.Write ".txtPortNo.value = """ & ConvSPChars(rs0("port_no")) & """" & vbCRLF
	
	Response.Write ".txtConfigFNm.value = """ & ConvSPChars(rs0("config_file_nm")) & """" & vbCRLF
	Response.Write ".txtConfigSNm.value = """ & ConvSPChars(rs0("config_step_nm")) & """" & vbCRLF
	Response.Write ".txtUrl.value = """ & ConvSPChars(rs0("url")) & """" & vbCRLF
	Response.Write ".txtEmail.value = """ & ConvSPChars(rs0("e_mail_id")) & """" & vbCRLF
	Response.Write ".txtLoginId.value = """ & ConvSPChars(rs0("login_id")) & """" & vbCRLF
	Response.Write ".txtLoginPwd.value = """ & ConvSPChars(rs0("login_pwd")) & """" & vbCRLF
	Response.Write ".txtRemark.value = """ & ConvSPChars(rs0("remark")) & """" & vbCRLF	

	Response.Write "parent.DbQueryOk" & vbCRLF
Response.Write "End With" & vbCRLF
Response.Write "</Script>" & vbCRLF
Response.End			
%>										