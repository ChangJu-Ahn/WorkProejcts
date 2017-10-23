<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List onhand stock detail
'*  3. Program ID           : I2161mb1.asp
'*  4. Program Name         : 실사선별 품목 조회 
'*  5. Program Desc         : 현재 창고에 있는 실사선별 할 품목의 상세정보를 조회한다.
'*  6. Comproxy List        : PI3S020.cILstOnhandStkDtl
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2000/04/03
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Jung Je Ahn
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf()

Err.Clear
On Error Resume Next
Call HideStatusWnd 
	
Dim pPI3S020										

Const C_SHEETMAXROWS_D = 100

Dim	I1_next_i_onhand_stock_tracking_no
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
    Const I303_EG1_E6_good_on_hand_qty = 19   

Dim StrNextKey1	
Dim StrNextKey2	
Dim StrNextKey3
Dim StrNextKey4
Dim lgStrKey1	
Dim lgStrKey2	
Dim lgStrKey3
Dim lgStrKey4

Dim LngMaxRow	
Dim LngRow
Dim strData
Dim PvArr
Dim GroupCount
	

	lgStrKey1 = Request("txthItemCd")
	lgStrKey2 = Request("txtTrackingNo2")
	lgStrKey3 = Request("txtLotNo")

	if Request("txtLotSubNo") <> "" then
		lgStrKey4 = CLng(Request("txtLotSubNo"))
	else
		lgStrKey4 = 0
	end if
	
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I7_b_plant_plant_cd							= Request("txtPlantCd")
    I5_b_storage_location_sl_cd					= Request("txtSLCd")
    I2_i_onhand_stock_tracking_no				= Request("txtTrackingNo")
    I8_i_onhand_stock_detail(I303_I8_lot_no)    = Request("txtLotNo")
    I6_b_item_item_cd							= Request("txtItemCd")
    I10_qty_flag(I303_I10_qty_flag)				= Request("txtFlag")
    
    If Trim(Request("txtValid")) <> "" Then
		I10_qty_flag(I303_I10_valid_flag) = Request("txtValid")    
	Else
		I10_qty_flag(I303_I10_valid_flag) = "N"
	End If
    
    If Request("txtLotSubNo") <> "" then
		I8_i_onhand_stock_detail(I303_I8_lot_sub_no) = Request("txtLotSubNo")
	End if
    
    If lgStrKey1 <> "" then
		I6_b_item_item_cd                           = lgStrKey1
        I1_next_i_onhand_stock_tracking_no			= lgStrKey2
        I8_i_onhand_stock_detail(I303_I8_lot_no)    = lgStrKey3
		
		If lgStrKey4 <> "" then
			I8_i_onhand_stock_detail(I303_I8_lot_sub_no) = lgStrKey4
		End if
	
	End if
	
    I4_query_type_action_entry = "NY"   'N:Item Y:Storage Location 창고별 품목재고정보 
    I3_ief_supplied_select_char = "P"   '신규추가:Block처리안된 항목만 Query
    

      Set pPI3S020 = Server.CreateObject("PI3S020.cILstOnhandStkDtl")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End											
	End If    

    '-----------------------
    'Com Action Area
    '-----------------------
	Call pPI3S020.I_LIST_ONHAND_STOCK_DETAIL(gStrGlobalCollection, C_SHEETMAXROWS_D, _
											I1_next_i_onhand_stock_tracking_no, I2_i_onhand_stock_tracking_no, _
											I3_ief_supplied_select_char,		I4_query_type_action_entry, _
											I5_b_storage_location_sl_cd,		I6_b_item_item_cd, _
											I7_b_plant_plant_cd,				I8_i_onhand_stock_detail, _
											I9_b_unit_of_measure,				I10_qty_flag, _
											E1_next_i_onhand_stock_tracking_no,	E2_next_b_item_item_cd, _
											E3_next_b_storage_location_sl_cd,	E4_next_i_onhand_stock_detail, _
											EG1_export_group)
 	
    
    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI3S020 = Nothing									
		Response.End											
	End If

  	Set pPI3S020 = Nothing											

	if isEmpty(EG1_export_group) then
		Response.End
	end if 
	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	GroupCount = ubound(EG1_export_group,1)
	
	ReDim PvArr(GroupCount)

	For LngRow = 0 To GroupCount

		strData = Chr(11) & "0" & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E3_item_cd)) & _
    			  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E3_item_nm)) & _
    			  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E3_spec)) & _
    			  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E1_tracking_no)) & _
    			  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E4_lot_no)) & _
    			  Chr(11) & ConvSPChars(EG1_export_group(LngRow,I303_EG1_E4_lot_sub_no)) & _
				  Chr(11) & LngMaxRow + LngRow & Chr(11)

        PvArr(LngRow) = strData
		    
	Next
	
    If EG1_export_group(GroupCount,I303_EG1_E3_item_cd)		= E2_next_b_item_item_cd and _
       EG1_export_group(GroupCount,I303_EG1_E1_tracking_no) = E1_next_i_onhand_stock_tracking_no and _
       EG1_export_group(GroupCount,I303_EG1_E4_lot_no)		= E4_next_i_onhand_stock_detail(I303_E5_lot_no) and _
       EG1_export_group(GroupCount,I303_EG1_E4_lot_sub_no)	= E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no) then
    		StrNextKey1 = ""
    		StrNextKey2 = ""
    		StrNextKey3 = ""
    		StrNextKey4 = ""
	Else
			StrNextKey1 = E2_next_b_item_item_cd
			StrNextKey2 = E1_next_i_onhand_stock_tracking_no
			StrNextKey3 = E4_next_i_onhand_stock_detail(I303_E5_lot_no)
			StrNextKey4 = E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no) 
	End If 

	
	strData = Join(PvArr, Chr(12)) & Chr(12)
    
    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " Dim i		 "               & vbCr
    Response.Write " With Parent "               & vbCr

	Response.Write "    .lgStrKey1  = """ & ConvSPChars(StrNextKey1) & """" & vbCr  
    Response.Write "    .lgStrKey2  = """ & ConvSPChars(StrNextKey2) & """" & vbCr  
	Response.Write "    .lgStrKey3  = """ & ConvSPChars(StrNextKey3) & """" & vbCr  
    Response.Write "    .lgStrKey4  = """ & ConvSPChars(StrNextKey4) & """" & vbCr  
	
	Response.Write "	.frm1.txthItemCd.value  = """ & ConvSPChars(EG1_export_group(GroupCount, I303_EG1_E3_item_cd)) & """" & vbCr  	   	  
  	Response.Write "	.frm1.txthPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr
  	Response.Write "	.frm1.txthSLCd.value    = """ & ConvSPChars(Request("txtSLCd")) & """" & vbCr

	Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr
	
    Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrKey1 <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	
	Response.Write "End with	" & vbCr
    Response.Write "</Script>      " & vbCr   

	Response.End     

%>



