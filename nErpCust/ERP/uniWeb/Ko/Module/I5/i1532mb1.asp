<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : Vendor Managed Inventory
'*  3. Program ID           : i1531mb1.asp
'*  4. Program Name         : List Production Order Header for VMI
'*  5. Program Desc         : List Production Order Header for VMI
'*  6. Modified date(First) : 2003-01-07
'*  7. Modified date(Last)  : 2003-04-28
'*  8. Modifier (First)     : Ahn, Jung Je
'*  9. Modifier (Last)      : Ahn, Jung Je
'* 10. Comment              : 
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE","MB")

Err.Clear
On Error Resume Next													
Call HideStatusWnd 
Dim pPI5G240												

Dim strData
Dim RemainQty
Dim NeedQty
Dim PvArr
Dim GroupCount
Dim LngMaxRow		
Dim LngRow    

Const C_SHEETMAXROWS_D = 100
    '-----------------------
    'IMPORTS View
    '-----------------------
	Dim I1_m_child_reserv
		Const I514_I1_plant_cd = 0
		Const I514_I1_bp_cd = 1
		Const I514_I1_item_cd = 2
		Const I514_I1_tracking_no = 3
		Const I514_I1_sl_cd = 4
	ReDim I1_m_child_reserv(I514_I1_sl_cd)

	'-----------------------
	'EXPORTS View
	'-----------------------
	Dim EG1_m_child_reserv
		Const I514_EG1_item_cd = 0
		Const I514_EG1_item_nm = 1
		Const I514_EG1_reqmt_qty = 2
		Const I514_EG1_reqmt_unit = 3
		Const I514_EG1_issue_qty = 4
		Const I514_EG1_tot_remain_qty = 5
		Const I514_EG1_good_onhand_qty = 6
		Const I514_EG1_sl_cd = 7
		Const I514_EG1_sl_nm = 8
		Const I514_EG1_tracking_no = 9
		Const I514_EG1_spec = 10

	Dim EG2_next_list
		Const I514_EG2_item_cd = 0
		Const I514_EG2_tracking_no = 1
		Const I514_EG2_sl_cd = 2

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_m_child_reserv(I514_I1_plant_cd)	= Request("txtPlantCd")
    I1_m_child_reserv(I514_I1_bp_cd)	= Request("txtSBPCd")
   
    If Request("lgStrPrevKey11") <> "" Then
		I1_m_child_reserv(I514_I1_item_cd)		= Request("lgStrPrevKey11")
		I1_m_child_reserv(I514_I1_tracking_no)	= Request("lgStrPrevKey12")
		I1_m_child_reserv(I514_I1_sl_cd)		= Request("lgStrPrevKey13")
    Else
        I1_m_child_reserv(I514_I1_item_cd)	= Request("txtItemCd")
    End If

	Set pPI5G240 = Server.CreateObject("PI5G240.cIVMIListSubcontItem")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    
	
	Call pPI5G240.I_VMI_LIST_SUNCONT_ITEM(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_m_child_reserv, _
										EG1_m_child_reserv, _
										EG2_next_list)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI5G240 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Set parent.gActiveElement = parent.document.activeElement" & vbCrLF
		Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If

	
	Set pPI5G240 = Nothing
	
	if isEmpty(EG1_m_child_reserv) then
		Response.End													
	end if
	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	GroupCount = ubound(EG1_m_child_reserv,1)
	
	ReDim PvArr(GroupCount)
	
	For LngRow = 0 To GroupCount

		RemainQty = CDbl(EG1_m_child_reserv(LngRow,I514_EG1_reqmt_qty)) - CDbl(EG1_m_child_reserv(LngRow,I514_EG1_issue_qty)) '출고잔량=필요수량-기출고수량 
		NeedQty   = CDbl(EG1_m_child_reserv(LngRow,I514_EG1_tot_remain_qty)) - CDbl(EG1_m_child_reserv(LngRow,I514_EG1_good_onhand_qty))                                     '입고필요량=출고잔량-재고수량 
		If NeedQty < 0 Then NeedQty = 0
		
		strData = Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_item_cd)) & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow ,I514_EG1_item_nm)) & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_tracking_no)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_m_child_reserv(LngRow, I514_EG1_reqmt_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_reqmt_unit)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_m_child_reserv(LngRow, I514_EG1_issue_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & UniConvNumberDBToCompany(RemainQty, ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _   
				  Chr(11) & UniConvNumberDBToCompany(EG1_m_child_reserv(LngRow, I514_EG1_good_onhand_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & UniConvNumberDBToCompany(NeedQty, ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _   
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_sl_cd)) & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_sl_nm)) & _
				  Chr(11) & ConvSPChars(EG1_m_child_reserv(LngRow, I514_EG1_spec)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 

        PvArr(LngRow) = strData

	Next

	strData = Join(PvArr, "")

	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.txthPlantCd.value			= """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr
	Response.Write "	.frm1.txthSBPCd.value = """ & ConvSPChars(Request("txtSBPCd")) & """" & vbCr
	Response.Write "	.frm1.txthItemCd.value          = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCr

	Response.Write "	.ggoSpread.Source	= .frm1.vspdData1 "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr

	Response.Write "    .lgStrPrevKey11  = """ & ConvSPChars(EG2_next_list(I514_EG2_item_cd)) & """" & vbCr  
    Response.Write "    .lgStrPrevKey12  = """ & ConvSPChars(EG2_next_list(I514_EG2_tracking_no)) & """" & vbCr  
	Response.Write "    .lgStrPrevKey13  = """ & ConvSPChars(EG2_next_list(I514_EG2_sl_cd)) & """" & vbCr  
	
	Response.Write "	If .frm1.vspdData1.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData1, 0) And .lgStrPrevKey11 <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End     	
%>
