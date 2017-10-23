<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2340mb4.asp
'*  4. Program Name			: Execute MRP&½ÂÀÎ in batch
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP2G101.cPExecMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/12/16
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment		:
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3


Dim strStatus
Dim strRunNo

	
	Redim UNISqlId(3)
	Redim UNIValue(3, 1)
	
	UNISqlId(0) = "185000saa"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "p2210mb2"
	UNISqlId(3) = "p2210mb2"

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNIValue(2, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(2, 1) = FilterVar(UniConvDate(Request("txtFixExecFromDt")), "''", "S")

	UNIValue(3, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(3, 1) = FilterVar(UniConvDate(Request("txtFixExecToDt")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	ELSE
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing				
		Response.End
	End If

    strStatus = rs0("status")
    strRunNo = right("000000000000" & Cstr(Clng(rs0("run_no")) + 1), 12)

    If strStatus = "1" Then	
		Call DisplayMsgBox("187731", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing         
		Response.End
	ElseIF  strStatus = "2" Or strStatus = "3" Then   
		Call DisplayMsgBox("187732", vbInformation, "", "", I_MKSCRIPT)   
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
    End If
    
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing
	
	If (rs2.EOF And rs2.BOF) Or (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("180200", vbInformation, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs2 = Nothing
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing         
		Response.End 
	End If
	
	rs2.Close
	Set rs2 = Nothing
	rs3.Close
	Set rs3 = Nothing

'--------------------------------------------------------------------------------------------------------------------									

    Dim pPP2G101	
	Dim FixExecFromDt
	Dim FixExecToDt
	Dim PlanExecToDt
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
    '-----------------------
    'Data manipulate area
    '-----------------------
    FixExecFromDt = UniConvDateToYYYYMMDD(Request("txtFixExecFromDt"),gDateFormat,"")
    FixExecToDt = UniConvDateToYYYYMMDD(Request("txtFixExecToDt"),gDateFormat,"")
    PlanExecToDt = UniConvDateToYYYYMMDD(Request("txtPlanExecToDt"),gDateFormat,"")
    
    I1_mrp_parameter(P202_I1_plant_cd) = UCase(Trim(Request("txtPlantCd")))	    
    I1_mrp_parameter(P202_I1_current_date) = FixExecFromDt
    I1_mrp_parameter(P202_I1_plan_date) =    FixExecToDt     
    I1_mrp_parameter(P202_I1_open_date) =    PlanExecToDt                                                                                                                                                                  
    I1_mrp_parameter(P202_I1_flag) = "B"

    If Request("rdoSafeInvFlg") = "Y" Then
         I1_mrp_parameter(P202_I1_safe_flg)  = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_safe_flg)  = "N"
    End If

    If Request("rdoAvailInvFlg") = "Y" Then
         I1_mrp_parameter(P202_I1_inv_flg)  = "Y"
    Else
    	 I1_mrp_parameter(P202_I1_inv_flg)  = "N"
    End If

    I1_mrp_parameter(P202_I1_idep_flg) = "Y"
    I1_mrp_parameter(P202_I1_option_flg) = "M"
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
    I1_mrp_parameter(P202_I1_mpsscope) = ""

	'-----------------------
    'Com Action Area
    '-----------------------
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
'----------------------------------------------------------------------------------------------------------------------

	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "185000saa"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtMRPHisNo.value = """ & ConvSPChars(rs0("run_no")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
    
	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)    
%>
