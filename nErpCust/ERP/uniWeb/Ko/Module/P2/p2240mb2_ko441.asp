<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->

<% Call LoadBasisGlobalInf
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2210mb2.asp
'*  4. Program Name			: MPS 일괄 생성 
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP2G101.cPExecMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/06/18
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************
	
Call HideStatusWnd

On Error Resume Next									

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+ 1. Pre Check : MPS 및 MRP 상태 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4

Dim strStatus
Dim strStatus_mrp
Dim dtf, ptf, PlanDt

	If Request("rdoStartDtFlg") = "N" And UniConvDate(Request("txtPlanDt")) <= UniConvDate(Request("txtPTF")) Then
	   Call DisplayMsgBox("183120", vbInformation, "", "", I_MKSCRIPT)
	   Response.End
	ElseIf Request("rdoStartDtFlg") = "Y" And UniConvDate(Request("txtPlanDt")) <= UniConvDate(Request("txtDTF")) Then
	   Call DisplayMsgBox("183121", vbInformation, "", "", I_MKSCRIPT)
	   Response.End   
	End IF

	Redim UNISqlId(4)
	Redim UNIValue(4, 1)
	
	UNISqlId(0) = "184000sab"			'Get Last MPS Run Status
	UNISqlId(1) = "185000saa"			'Get Last Mrp Run Status
	UNISqlId(2) = "184000saa"			'Plant Name 조회 
	UNISqlId(3) = "p2210mb2"
	UNISqlId(4) = "p2210mb2"

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")

	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(1, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNIValue(2, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	dtf =  UNIDateAdd("d",1,Request("txtDTF"), gDateFormat)
    dtf = UniConvDate(dtf)
    
    ptf =  UNIDateAdd("d",1,Request("txtPTF"), gDateFormat)
    ptf = UniConvDate(ptf)
    
    PlanDt = UniConvDate(Request("txtPlanDt"))

	UNIValue(3, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
		
	IF  Request("rdoStartDtFlg") = "Y" Then
		UNIValue(3, 1) = FilterVar(dtf, "''", "S")
    Else
		UNIValue(3, 1) = FilterVar(ptf, "''", "S")
    End If
        
	UNIValue(4, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(4, 1) = FilterVar(PlanDt, "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
	
	'---------------------------
	'공장이 없으면 종료 
	'---------------------------
	If rs2.EOF And rs2.BOF Then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing  		
		rs2.Close
		Set rs2 = Nothing		
		Set ADF = Nothing 
		Response.End													
	End If    
	
	'---------------------------
	'MPS 일괄 생성 상태이면 종료 
	'---------------------------
    If Not(rs0.EOF And rs0.BOF) Then
		strStatus = rs0("status")
		If strStatus = "1" Then											
		    Call DisplayMsgBox("187730", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		End If
	End If
    
    '----------------------------------------------
	'MPS 전개 상태이거나 승인 상태이면 종료 
	'----------------------------------------------
    If Not(rs1.EOF And rs1.BOF) Then
        strStatus_mrp = rs1("status")
		If strStatus_mrp = "1" Then	
			Call DisplayMsgBox("187731", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		ElseIf  strStatus_mrp = "2" Or strStatus_mrp = "3" Then
			Call DisplayMsgBox("187732", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		End If
	End IF   
    
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing	
	
	If (rs3.EOF And rs3.BOF) Or (rs4.EOF And rs4.BOF) Then
		Call DisplayMsgBox("180200", vbInformation, "", "", I_MKSCRIPT)
		rs3.Close
		Set rs3 = Nothing
		rs4.Close
		Set rs4 = Nothing
		Set ADF = Nothing         
		Response.End 
	End If
	
	rs3.Close
	Set rs3 = Nothing
	rs4.Close
	Set rs4 = Nothing

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+ 2. MPS 일괄생성 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Dim pPP2G101
    Dim mpsscope
    Dim I1_mrp_parameter

    Err.Clear
	
	Const P202_I1_plant_cd = 0
	Const P202_I1_current_date = 1
	Const P202_I1_plan_date = 2
	Const P202_I1_open_date = 3
	Const P202_I1_flag = 4
	Const P202_I1_safe_flg = 5
	Const P202_I1_inv_flg = 6
	Const P202_I1_idep_flg = 7
	Const P202_I1_option_flg = 8
	Const P202_I1_item_cd = 9
	Const P202_I1_warning_flg = 10
	Const P202_I1_order_no = 11
	Const P202_I1_codr_flg = 12
	Const P202_I1_net_flg = 13
	Const P202_I1_pack_flg = 14
	Const P202_I1_scrap = 15
	Const P202_I1_forward = 16
	Const P202_I1_mpsscope = 17
    
    Redim I1_mrp_parameter(P202_I1_mpsscope)

	dtf =  UNIDateAdd("d",1,Request("txtDTF"), gDateFormat)
    dtf = UniConvDateToYYYYMMDD(dtf,gDateFormat,"")
    
    ptf =  UNIDateAdd("d",1,Request("txtPTF"), gDateFormat)
    ptf = UniConvDateToYYYYMMDD(ptf,gDateFormat,"")
    
    PlanDt = UniConvDateToYYYYMMDD(Request("txtPlanDt"),gDateFormat,"")
    
    I1_mrp_parameter(P202_I1_plant_cd) = UCase(Trim(Request("txtPlantCd")))
    
    IF  Request("rdoStartDtFlg") = "Y" Then
		I1_mrp_parameter(P202_I1_current_date) = dtf
    Else
		I1_mrp_parameter(P202_I1_current_date) = ptf
    End If

	I1_mrp_parameter(P202_I1_plan_date) = ptf
	I1_mrp_parameter(P202_I1_open_date) = PlanDt
    
    IF  Request("rdoStartDtFlg") = "Y" Then
        I1_mrp_parameter(P202_I1_flag) = "1"
    Else
        I1_mrp_parameter(P202_I1_flag) = "2"
    End If

    If Request("rdoSafeInvFlg") = "Y" Then
		I1_mrp_parameter(P202_I1_safe_flg)  = "Y"
    Else
		I1_mrp_parameter(P202_I1_safe_flg)  = "N"
    End If

    If Request("rdoAvailInvFlg") = "Y" Then
		I1_mrp_parameter(P202_I1_inv_flg) = "Y"
    Else
		I1_mrp_parameter(P202_I1_inv_flg) = "N"
    End If

	I1_mrp_parameter(P202_I1_idep_flg) = "Y"
	I1_mrp_parameter(P202_I1_option_flg) = "P"
	I1_mrp_parameter(P202_I1_item_cd) = "%"
	
	I1_mrp_parameter(P202_I1_warning_flg) = "N"
	I1_mrp_parameter(P202_I1_order_no) = ""
	I1_mrp_parameter(P202_I1_codr_flg) = "Y"
    
    If Request("rdoAvailInvFlg") = "Y" Then
		I1_mrp_parameter(P202_I1_net_flg)  = "Y"
    Else
		I1_mrp_parameter(P202_I1_net_flg)  = "N"
    End If

    I1_mrp_parameter(P202_I1_pack_flg) = "N"
    I1_mrp_parameter(P202_I1_scrap) = ""
    I1_mrp_parameter(P202_I1_forward) = ""
    
    mpsscope = "%"
    If Request("rdoMinFlg") = "Y" Then
       mpsscope = mpsscope & "Y"
    Else
       mpsscope = mpsscope & "N"
    End If
    
    If Request("rdoMaxFlg") = "Y" Then
       mpsscope = mpsscope & "Y"
    Else
       mpsscope = mpsscope & "N"
    End If

    If Request("rdoRoundFlg") = "Y" Then
       mpsscope = mpsscope & "Y"
    Else
       mpsscope = mpsscope & "N"
    End If
    
    IF  Request("rdoStartDtFlg") = "Y" Then
        mpsscope = mpsscope & "D"
    ELSE
        mpsscope = mpsscope & "P"
    END IF
    
	I1_mrp_parameter(P202_I1_mpsscope) = mpsscope


	Set pPP2G101 = Server.CreateObject("PP2G101.cPExecMrpSvr")
	    
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G101 = Nothing		
		Response.End
	End If
	
	Call pPP2G101.P_EXEC_MRP_SVR(gStrGlobalCollection, I1_mrp_parameter)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G101 = Nothing
		Response.End
	End If

	Set pPP2G101 = Nothing

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+ 3. MPS Run Number 조회 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "184000sab"
	

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtMPSHistoryNo.value = """ & ConvSPChars(rs0("mps_history_no")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
	rs0.Close
	Set rs0 = Nothing
	
    Set ADF = Nothing
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)    
%>
