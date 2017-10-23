<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List onhand stock detail
'*  3. Program ID           : I2511rb2.asp
'*  4. Program Name         : 재고 상세 조회 
'*  5. Program Desc         : LOT에 해당하는 있는 품목의 상세정보를 조회한다.
'*  6. Comproxy List        : 
'                             
'                             
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2000/04/03
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Nam hoon kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 

Err.Clear
On Error Resume Next
Dim i22128           
Dim strData
Dim strMode          

Dim StrNextKey  
Dim StrNextSubKey  
Dim LngMaxRow  
Dim LngRow
Dim GroupCount  
Dim PvArr        

	Const C_SHEETMAXROWS_D = 100
    
'-----------------------
'IMPORTS View
'-----------------------
Dim I1_next_i_onhand_stock_tracking_no
Dim I2_i_onhand_stock_tracking_no
Dim I3_ief_supplied_select_char
Dim I4_query_type_action_entry
Dim I5_b_storage_location_sl_cd
Dim I6_b_item_item_cd
Dim I7_b_plant_plant_cd
Dim I8_i_onhand_stock_detail
	Const I303_I8_lot_no = 0
	Const I303_I8_lot_sub_no = 1
ReDim I8_i_onhand_stock_detail(I303_I8_lot_sub_no)
Dim I9_b_unit_of_measure
	Const I303_I9_unit = 0
	Const I303_I9_unit_nm = 1 
ReDim I9_b_unit_of_measure(I303_I9_unit_nm)

Dim I10_qty_flag
		Const I303_I10_qty_flag = 0
		Const I303_I10_valid_flag = 1
ReDim I10_qty_flag(I303_I10_valid_flag)
'-----------------------
'EXPORTS View
'-----------------------
Dim E1_b_plant
	Const I303_E1_plant_cd = 0
	Const I303_E1_plant_nm = 1
Dim E1_next_i_onhand_stock_tracking_no
Dim E2_next_b_item_item_cd
Dim E3_next_b_storage_location_sl_cd
Dim E4_next_i_onhand_stock_detail
	Const I303_E5_lot_no = 0
	Const I303_E5_lot_sub_no = 1
Dim EG1_export_group
	Const I303_EG1_E1_tracking_no = 0
	Const I303_EG1_E2_sl_cd = 1
	Const I303_EG1_E2_sl_nm = 2
	Const I303_EG1_E3_item_cd = 3
	Const I303_EG1_E3_item_nm = 4
	Const I303_EG1_E3_spec = 5
	Const I303_EG1_E3_basic_unit = 6
	Const I303_EG1_E4_lot_no = 7
	Const I303_EG1_E4_lot_sub_no = 8
	Const I303_EG1_E4_good_on_hand_qty = 9
	Const I303_EG1_E4_bad_on_hand_qty = 10
	Const I303_EG1_E4_stk_on_insp_qty = 11
	Const I303_EG1_E4_stk_on_trns_qty = 12
	Const I303_EG1_E4_picking_qty = 13
	Const I303_EG1_E4_prev_good_qty = 14
	Const I303_EG1_E4_prev_bad_qty = 15
	Const I303_EG1_E4_prev_stk_on_insp_qty = 16
	Const I303_EG1_E4_prev_stk_in_trns_qty = 17
	Const I303_EG1_E5_abc_flg = 18


     
	StrNextKey    = Request("lgStrPrevKey")
	StrNextSubKey = Request("lgStrPrevSubKey")
	 
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	I4_query_type_action_entry = "YN"   'Y:Storage Location N:Item LOT별품목재고정보 
	I6_b_item_item_cd  = Request("txtItem_Cd")
	I7_b_plant_plant_cd = Request("txtPlant_Cd")
	' I5_b_storage_location_sl_cd        = StrNextKey
	' I1_next_i_onhand_stock_tracking_no = StrNextSubKey
	 
	I8_i_onhand_stock_detail(I303_I8_lot_no)   = Request("txtLot_No")
	I8_i_onhand_stock_detail(I303_I8_lot_sub_no) = Request("txtLotSub_No")
	    
	if StrNextKey <> "" And StrNextSubKey <> "" then 
		I8_i_onhand_stock_detail(I303_I8_lot_no)   = StrNextKey
		I8_i_onhand_stock_detail(I303_I8_lot_sub_no) = StrNextSubKey
	end if    

	If CheckSYSTEMError(Err, True) = True Then
		Response.End            '☜: 비지니스 로직 처리를 종료함 
	End If   
	     
	    
	Set i22128 = Server.CreateObject("PI3S020.cILstOnhandStkDtl")     
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End            '☜: 비지니스 로직 처리를 종료함 
	End If    
	 
	Call i22128.I_LIST_ONHAND_STOCK_DETAIL(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_next_i_onhand_stock_tracking_no, _
										I2_i_onhand_stock_tracking_no, _
										I3_ief_supplied_select_char, _
										I4_query_type_action_entry, _
										I5_b_storage_location_sl_cd, _
										I6_b_item_item_cd, _
										I7_b_plant_plant_cd, _
										I8_i_onhand_stock_detail, _
										I9_b_unit_of_measure, _
										I10_qty_flag, _
										E1_next_i_onhand_stock_tracking_no, _
										E2_next_b_item_item_cd, _
										E3_next_b_storage_location_sl_cd, _
										E4_next_i_onhand_stock_detail, _
										EG1_export_group)
	                                    
	If CheckSYSTEMError(Err, True) = True Then
		Set i22128 = Nothing
		Response.End
	End If

	Set i22128 = Nothing    

	if IsEmpty(EG1_export_group) then
		Response.End
	end if
	 
	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows"))
	GroupCount = ubound(EG1_export_group,1)
	ReDim PvArr(ubound(EG1_export_group,1))
	     
	For LngRow = 0 To GroupCount
		strData = Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E2_sl_cd)) & _
		Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E2_sl_nm)) & _
		Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E1_tracking_no)) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_good_on_hand_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_bad_on_hand_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_stk_on_insp_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_stk_on_trns_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_good_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_bad_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_stk_on_insp_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_stk_in_trns_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
	PvArr(LngRow) = strData
	Next
	
	strData = Join(PvArr, "")
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If

	If   EG1_export_group(GroupCount, I303_EG1_E4_lot_no)      =      E4_next_i_onhand_stock_detail(I303_E5_lot_no) And _
	CInt(EG1_export_group(GroupCount, I303_EG1_E4_lot_sub_no)) = CInt(E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no)) then

		StrNextKey    = ""
		StrNextSubKey = ""
	else
		StrNextKey    = i22128.E4_next_i_onhand_stock_detail(I303_E5_lot_no)
		StrNextSubKey = E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no)     
	End If    

	Response.Write "<Script Language=vbscript> " & vbcr
	Response.Write "With parent "                & vbcr    
	  
	Response.Write " .txtItem_Nm.value = """ & ConvSPChars(i22128.ExportItemBItemItemNm(1)) & """ " & vbcr
	    
	Response.Write " .ggoSpread.Source = .vspdData " & vbcr 
	Response.Write " .ggoSpread.SSShowData """ & strData & """ " & vbcr
	   
	Response.Write " .lgStrPrevKey =    """ & ConvSPChars(StrNextKey)    & """ " & vbcr
	Response.Write " .lgStrPrevSubKey = """ & ConvSPChars(StrNextSubKey) & """ " & vbcr
	     
	Response.Write " If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbCr
	Response.Write "   .DbQuery "                                                            & vbCr
	Response.Write " Else "                                                                  & vbCr
	Response.Write "   .DbQueryOk "                                                          & vbCr
	Response.Write "    End If "                                                             & vbCr
	
	Response.Write "End With "       & vbcr
	Response.Write "</Script> "      & vbcr
%>


