<%@LANGUAGE = VBScript%>
<%Option Explicit
'======================================================================================================
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
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
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
Dim pPI5G210													'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strData
Dim RemainQty
Dim NeedQty

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow    
Dim PvArr
Dim GroupCount

Const C_SHEETMAXROWS_D = 100
    '-----------------------
    'IMPORTS View
    '-----------------------
	Dim I1_p_reservation
		Const I511_I1_plant_cd = 0
		Const I511_I1_item_cd = 1
		Const I511_I1_tracking_no = 2
		Const I511_I1_sl_cd = 3
		Const I511_I1_req_start_dt = 4
		Const I511_I1_req_end_dt = 5
		Const I511_I1_start_prod_order_no = 6
		Const I511_I1_end_prod_order_no = 7
	ReDim I1_p_reservation(I511_I1_end_prod_order_no)
	
	'-----------------------
	'EXPORTS View
	'-----------------------
	Dim EG1_p_reservation
		Const I511_EG1_item_cd = 0
		Const I511_EG1_item_nm = 1
		Const I511_EG1_resvr_qty = 2
		Const I511_EG1_basic_unit = 3
		Const I511_EG1_issued_qty = 4
		Const I511_EG1_tot_remain_qty = 5
		Const I511_EG1_good_onhand_qty = 6
		Const I511_EG1_sl_cd = 7
		Const I511_EG1_sl_nm = 8
		Const I511_EG1_tracking_no = 9
		Const I511_EG1_spec = 10

	Dim EG2_next_list
		Const I511_EG2_item_cd = 0
		Const I511_EG2_tracking_no = 1
		Const I511_EG2_sl_cd = 2

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_p_reservation(I511_I1_plant_cd)				= Request("txtPlantCd")
    I1_p_reservation(I511_I1_req_start_dt)			= UNIConvDate(Request("txtReqStartDt"))
    I1_p_reservation(I511_I1_req_end_dt)			= UNIConvDate(Request("txtReqEndDt"))
    I1_p_reservation(I511_I1_start_prod_order_no)	= Request("txtProdOrderNoFrom")
    I1_p_reservation(I511_I1_end_prod_order_no)     = Request("txtProdOrderNoTo")
   
    If Request("lgStrPrevKey11") <> "" Then
		I1_p_reservation(I511_I1_item_cd)		= Request("lgStrPrevKey11")
		I1_p_reservation(I511_I1_tracking_no)	= Request("lgStrPrevKey12")
		I1_p_reservation(I511_I1_sl_cd)			= Request("lgStrPrevKey13")
    Else
        I1_p_reservation(I511_I1_item_cd)	= Request("txtItemCd")
    End If

	Set pPI5G210 = Server.CreateObject("PI5G210.cIVMIListProdItem")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												'☜: 비지니스 로직 처리를 종료함 
	End If    
	
	Call pPI5G210.I_VMI_LIST_PROD_ITEM(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_p_reservation, _
										EG1_p_reservation, _
										EG2_next_list)


    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI5G210 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Set parent.gActiveElement = parent.document.activeElement" & vbCrLF
		Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
		'Call ServerMesgBox("Call 실패 ERR_number :" & Err.number & " ERR_description :" & Err.description , vbCritical, I_MKSCRIPT)  
	End If

	
	Set pPI5G210 = Nothing
	
	if isEmpty(EG1_p_reservation) then
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	end if
	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1

	GroupCount = ubound(EG1_p_reservation,1)
	
	ReDim PvArr(GroupCount)
	
	For LngRow = 0 To GroupCount
		
		RemainQty = CDbl(EG1_p_reservation(LngRow,I511_EG1_resvr_qty)) - CDbl(EG1_p_reservation(LngRow,I511_EG1_issued_qty)) '출고잔량=필요수량-기출고수량 
		NeedQty   = CDbl(EG1_p_reservation(LngRow,I511_EG1_tot_remain_qty)) - CDbl(EG1_p_reservation(LngRow,I511_EG1_good_onhand_qty))                                     '입고필요량=출고잔량-재고수량 
		If NeedQty < 0 Then NeedQty = 0
		
		strData = Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_item_cd)) & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow ,I511_EG1_item_nm)) & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_tracking_no)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_p_reservation(LngRow, I511_EG1_resvr_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_basic_unit)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_p_reservation(LngRow, I511_EG1_issued_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & UniConvNumberDBToCompany(RemainQty, ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _   
				  Chr(11) & UniConvNumberDBToCompany(EG1_p_reservation(LngRow, I511_EG1_good_onhand_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & UniConvNumberDBToCompany(NeedQty, ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _   
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_sl_cd)) & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_sl_nm)) & _
				  Chr(11) & ConvSPChars(EG1_p_reservation(LngRow, I511_EG1_spec)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
	
        PvArr(LngRow) = strData

	Next

	strData = Join(PvArr, "")

	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.txthPlantCd.value			= """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr
	Response.Write "	.frm1.txthProdOrderNoFrom.value = """ & ConvSPChars(Request("txtProdOrderNoFrom")) & """" & vbCr
	Response.Write "	.frm1.txthProdOrderNoTo.value   = """ & ConvSPChars(Request("txtProdOrderNoTo")) & """" & vbCr
	Response.Write "	.frm1.txthItemCd.value          = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCr
	Response.Write "	.frm1.txthReqStartDt.value      = """ & ConvSPChars(Request("txtReqStartDt")) & """" & vbCr
	Response.Write "	.frm1.txthReqEndDt.value        = """ & ConvSPChars(Request("txtReqEndDt")) & """" & vbCr

	Response.Write "	.ggoSpread.Source	= .frm1.vspdData1 "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr

	Response.Write "    .lgStrPrevKey11  = """ & ConvSPChars(EG2_next_list(I511_EG2_item_cd)) & """" & vbCr  
    Response.Write "    .lgStrPrevKey12  = """ & ConvSPChars(EG2_next_list(I511_EG2_tracking_no)) & """" & vbCr  
	Response.Write "    .lgStrPrevKey13  = """ & ConvSPChars(EG2_next_list(I511_EG2_sl_cd)) & """" & vbCr  
	
	Response.Write "	If .frm1.vspdData1.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData1, 0) And .lgStrPrevKey11 <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr

	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End     
	
%>