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
'*  3. Program ID			: p2348mb1.asp
'*  4. Program Name			: Query Available Inventory
'*  5. Program Desc			: 
'*  6. Comproxy List		: PP2G101.cPExecMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/06/18
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment		:
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")  

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4

Dim lgStrPrevKey	' ÀÌÀü °ª 
Dim i
Dim f_ss_qty

On Error Resume Next

Dim strTrackingNo

	lgStrPrevKey = Request("lgStrPrevKey")
	
	
'--------------------------------------------------------------------------------------------------------------------	
    Dim pPP2G101
    Dim strcurrentdate
    Dim CurrentDate
    Dim PlanDate
    Dim OpenDate
    
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
	
    CurrentDate = UniConvDateToYYYYMMDD(GetSvrDate,gServerDateFormat,"")
    PlanDate = UNIDateAdd("YYYY", 1, GetSvrDate, gServerDateFormat)
    PlanDate = UniConvDateToYYYYMMDD(PlanDate,gServerDateFormat,"")
    OpenDate = PlanDate
    
    I1_mrp_parameter(P202_I1_plant_cd) = UCase(Trim(Request("txtPlantCd")))
    I1_mrp_parameter(P202_I1_current_date) = CurrentDate
    I1_mrp_parameter(P202_I1_plan_date) = PlanDate
    I1_mrp_parameter(P202_I1_open_date) = OpenDate
                                        
    I1_mrp_parameter(P202_I1_flag) = ""
    I1_mrp_parameter(P202_I1_safe_flg) = "N"
    I1_mrp_parameter(P202_I1_inv_flg) = "Y"
    I1_mrp_parameter(P202_I1_idep_flg) = "N"
    I1_mrp_parameter(P202_I1_option_flg) = "I"
    I1_mrp_parameter(P202_I1_item_cd) = UCase(Trim(Request("txtItemCd")))

    I1_mrp_parameter(P202_I1_warning_flg) = "N"
    I1_mrp_parameter(P202_I1_order_no) = ""
    I1_mrp_parameter(P202_I1_codr_flg) = "N"
    I1_mrp_parameter(P202_I1_net_flg) = "Y"
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
    	
	Redim UNISqlId(4)
	Redim UNIValue(4, 2)
	
	UNISqlId(0) = "185800saa"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "180000saf"
	
	IF Request("txtTrackingNo") = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	END IF

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(0, 2) = strTrackingNo
	
	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(2, 0) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(4, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If

	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("item_nm")) & """" & vbCrLf
			Response.Write "parent.frm1.txtBasicUnit.value = """ & ConvSPChars(rs2("basic_unit")) & """" & vbCrLf
			Response.Write "parent.frm1.txtItemSpec.value = """ & ConvSPChars(rs2("spec")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If
	
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF AND rs3.BOF) Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			
			Response.End
		End If
	End If

	If Not(rs4.EOF AND rs4.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSsQty.value = """ & UniConvNumberDBToCompany(rs4("ss_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & """" & vbCrLf
			f_ss_qty = rs4("ss_qty")
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		rs2.Close
		Set rs2 = Nothing		
		rs3.Close
		Set rs3 = Nothing
		rs4.Close
		Set rs4 = Nothing	
		Set ADF = Nothing				
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim f_avl_qty
Dim f_onhand_qty
Dim arrVal
ReDim arrVal(0)
  	
With parent	
	LngMaxRow = .frm1.vspdData.MaxRows
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		
<%      IF i = 0 Then
%>			f_avl_qty = parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>") _
				+ parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>") _
				- parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>") _
				- parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(f_ss_qty,ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>")
			f_onhand_qty = parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>")
<%      ELSE
%>
			f_avl_qty = f_avl_qty + parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>") _
				+ parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>") _
				- parent.parent.uniCdbl("<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>")
<%      End If
%>
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt"))%>"
		strData = strData & Chr(11) & parent.parent.UNIFormatNumber(f_avl_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("sch_rcpt_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("resrv_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("on_hand_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		ReDim Preserve arrVal(<%=i%>)
		arrVal(<%=i%>) = strData
		
<%		
		rs0.MoveNext
	Next
%>
		.frm1.txtOnhandQty.value = parent.parent.UNIFormatNumber(f_onhand_qty, .ggQty.DecPoint,-2, 0,.ggQty.RndPolicy,.ggQty.RndUnit)

		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData Join(arrVal,"")
		
		.lgStrPrevKey = "<%=StrNextKey%>"

<%			
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		rs2.Close
		Set rs2 = Nothing
		rs3.Close
		Set rs3 = Nothing
		rs4.Close
		Set rs4 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing
%>
