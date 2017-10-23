<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4110mb3.asp
'*  4. Program Name			: ��ȹ����Ȯ�� 
'*  5. Program Desc			: Confirm Mrp
'*  6. Comproxy List		: PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/08/20
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call HideStatusWnd

On Error Resume Next									'��: 
'--------------------------------------------------------------------------------------------------------------------
Dim ADF													'ActiveX Data Factory ���� �������� 
Dim strRetMsg											'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0			'DBAgent Parameter ���� 
Dim lgStrPrevKey

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim strStatus

	lgStrPrevKey = Request("lgStrPrevKey")
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	UNISqlId(0) = "189702sae"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(0, 2) = FilterVar(UCase(Request("txtPlanOrderNO")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("187734", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing		
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If

    strStatus = rs0("confirm_flg")
    If strStatus = "Y" Then	'��: ������ ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("187743", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing         
		Response.End
    End If

	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing													'��: ActiveX Data Factory Object Nothing

'--------------------------------------------------------------------------------------------------------------------								
    Dim pPP2G152												'�� : �Է�/������ ComProxy Dll ��� ���� 
    Dim I1_plant_cd
    Dim I2_mrp_parameter
    
    Const P206_I2_plant_cd = 0    
    Const P206_I2_safe_flg = 1
    Const P206_I2_inv_flg = 2
    Const P206_I2_idep_flg = 3
    Const P206_I2_forward = 4
    Const P206_I2_mpsscope = 5
	
	Dim I3_select_char
	Dim I4_ord_no
	Dim txtSpread1, txtSpread2
	
	
	txtSpread1 = Request("txtSpread")
	txtSpread2 = Request("txtSpread2")
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_safe_flg) = "Y"
	I2_mrp_parameter(P206_I2_inv_flg) = "C"
	I2_mrp_parameter(P206_I2_idep_flg) = "S" 
	I2_mrp_parameter(P206_I2_forward) = UCase(Trim(Request("txtPlanOrderNo")))
	I2_mrp_parameter(P206_I2_mpsscope) = "" 

	I3_select_char = "S"
	
	I4_ord_no = UCase(Trim(Request("txtPlanOrderNo")))
	
	'I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Request("txtPlantCd"))
	'I2_mrp_parameter(P206_I2_safe_flg)	= "Y"
	'I2_mrp_parameter(P206_I2_inv_flg)	= "M"
	'I2_mrp_parameter(P206_I2_idep_flg)	= "S" 
	'I2_mrp_parameter(P206_I2_forward)	= UCase(Request("hMrpNo"))
	'I2_mrp_parameter(P206_I2_mpsscope)	= "" 

	'I3_select_char = "S"
	
    '-----------------------
    'Com Action Area
    '-----------------------
    
    Set pPP2G152 = Server.CreateObject("PP2G152.cPCnfmOrdExpSvr")
	    
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G152 = Nothing		
		Response.End
	End If
	
	Call pPP2G152.P_CONFIRM_ORD_EXP_SVR(gStrGlobalCollection, _
									I1_plant_cd, _
									I2_mrp_parameter, _
									I3_select_char, _
									I4_ord_no, _
									txtSpread1, _
									txtSpread2)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G152 = Nothing
		Response.End
	End If
	
	Set pPP2G152 = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "	Call Parent.ConfirmOk()  " & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
	
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
	
%>
