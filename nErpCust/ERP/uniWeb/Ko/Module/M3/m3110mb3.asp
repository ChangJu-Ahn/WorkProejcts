<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4110mb3.asp
'*  4. Program Name			: 계획오더확정 
'*  5. Program Desc			: Confirm Mrp
'*  6. Comproxy List		: PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/08/20
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
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

On Error Resume Next									'☜: 
'--------------------------------------------------------------------------------------------------------------------
Dim ADF													'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg											'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0			'DBAgent Parameter 선언 
Dim lgStrPrevKey

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim strStatus

	lgStrPrevKey = Request("lgStrPrevKey")
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	UNISqlId(0) = "189702sae"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = "" & FilterVar("C", "''", "S") & " "
	UNIValue(0, 2) = FilterVar(UCase(Request("txtPlanOrderNO")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("187734", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing		
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

    strStatus = rs0("confirm_flg")
    If strStatus = "Y" Then	'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("187743", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing         
		Response.End
    End If

	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing													'☜: ActiveX Data Factory Object Nothing

'--------------------------------------------------------------------------------------------------------------------								
    Dim pPP2G102												'☆ : 입력/수정용 ComProxy Dll 사용 변수 
    Dim I1_plant_cd
    Dim I2_mrp_parameter
    
    Const P206_I2_plant_cd = 0    
    Const P206_I2_safe_flg = 1
    Const P206_I2_inv_flg = 2
    Const P206_I2_idep_flg = 3
    Const P206_I2_forward = 4
    Const P206_I2_mpsscope = 5
	
	Dim I3_select_char
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_safe_flg) = "Y"
	I2_mrp_parameter(P206_I2_inv_flg) = "C"
	I2_mrp_parameter(P206_I2_idep_flg) = "M" 
	I2_mrp_parameter(P206_I2_forward) = UCase(Trim(Request("txtPlanOrderNo")))
	I2_mrp_parameter(P206_I2_mpsscope) = "" 

	I3_select_char = "M"
    '-----------------------
    'Com Action Area
    '-----------------------
    
    Set pPP2G102 = Server.CreateObject("PP2G102.cPCnfmMrpSvr")
	    
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G102 = Nothing		
		Response.End
	End If
	
	Call pPP2G102.P_CONFIRM_MRP_SRV(gStrGlobalCollection, _
									I1_plant_cd, _
									I2_mrp_parameter, _
									I3_select_char, _
									"")

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G102 = Nothing
		Response.End
	End If

	Set pPP2G102 = Nothing
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
%>
