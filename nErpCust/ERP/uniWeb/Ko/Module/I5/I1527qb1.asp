<%@ LANGUAGE=VBSCript%>
<% Option Explicit%> 
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock 
'*  2. Function Name        : 
'*  3. Program ID           : I1527qb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 현 재고 조회(VMI)
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2003/02/18
'*  8. Modified date(Last)  : 2003/02/18
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")   

Call HideStatusWnd 

On Error Resume Next
Err.Clear

Dim i22118           
Dim strMode          

Dim StrNextKey1  
Dim StrNextKey2  
Dim LngMaxRow  
Dim LngRow
Dim strData
Dim PvArr

Const C_SHEETMAXROWS_D = 500

 Dim import_i_onhand_stock_tracking_no   
 Dim import_b_item
    Const C_import_b_item_item_cd = 0
    Const C_import_b_item_basic_unit = 1          
 ReDim import_b_item(C_import_b_item_basic_unit)
 Dim import_b_storage_location_sl_cd  
 Dim import_b_plant_cd
 
 Dim Import_flag
	Const C_vmi_flag = 0
	Const C_valid_flag = 1
	Const C_qty_flag = 2
 ReDim Import_flag(C_qty_flag)              

 Dim export_next_i_onhand_stock          
 Dim export_next_b_item_cd                  
 Dim export_group                        
    Const C_export_group_export_b_item_by_plant_location = 0
    Const C_export_group_export_b_plant_plant_cd = 1
    Const C_export_group_export_b_plant_plant_nm = 2
    Const C_export_group_export_b_storage_location_sl_cd = 3
    Const C_export_group_export_b_storage_location_sl_type = 4
    Const C_export_group_export_b_storage_location_sl_nm = 5
    Const C_export_group_export_b_item_item_cd = 6
    Const C_export_group_export_b_item_item_nm = 7
    Const C_export_group_export_b_item_spec = 8
    Const C_export_group_export_b_item_item_acct = 9
    Const C_export_group_export_b_item_item_class = 10
    Const C_export_group_export_b_item_basic_unit = 11
    Const C_export_group_export_i_onhand_stock_tracking_no = 12
    Const C_export_group_export_i_onhand_stock_block_indicator = 13
    Const C_export_group_export_i_onhand_stock_good_on_hand_qty = 14
    Const C_export_group_export_i_onhand_stock_bad_on_hand_qty = 15
    Const C_export_group_export_i_onhand_stock_stk_on_insp_qty = 16
    Const C_export_group_export_i_onhand_stock_stk_in_trns_qty = 17
    Const C_export_group_export_i_onhand_stock_prev_good_qty = 18
    Const C_export_group_export_i_onhand_stock_prev_bad_qty = 19
    Const C_export_group_export_i_onhand_stock_prev_stk_on_insp_qty = 20
    Const C_export_group_export_i_onhand_stock_prev_stk_in_trns_qty = 21
    Const C_export_group_export_i_onhand_stock_curr_yr = 22
    Const C_export_group_export_i_onhand_stock_curr_mnth = 23
    Const C_export_group_export_i_onhand_stock_schd_rcpt_qty = 24
    Const C_export_group_export_i_onhand_stock_schd_issue_qty = 25
    Const C_export_group_export_i_onhand_stock_allocation_qty = 26
    Const C_export_group_export_i_onhand_stock_detail_picking_qty = 27
    
 
 StrNextKey1 = Request("lgStrPrevKey1")
 StrNextKey2 = Request("lgStrPrevKey2")
 
 import_b_storage_location_sl_cd					= Request("txtSL_Cd")
 import_b_item(C_import_b_item_item_cd)				= Request("txtItem_Cd")
 import_b_item(C_import_b_item_basic_unit)			= Request("txtinvunit") 
 import_b_plant_cd									= Request("txtPlant_Cd")
 Import_flag(C_vmi_flag)							= Request("txthUserFlag")
 Import_flag(C_valid_flag)							= Request("txtCheck")
 Import_flag(C_qty_flag)							= Request("txtQtyCheck")

 If StrNextKey1 <> "" then import_b_item(C_import_b_item_item_cd) = StrNextKey1  
 IF StrNextKey2 <> "" then import_i_onhand_stock_tracking_no = StrNextKey2 '""

 Set i22118 = Server.CreateObject("PI3G010.cIListOnhandStkSvr")     

 If CheckSYSTEMError(Err, True) = True Then
	Response.End           
 End If    

    Call i22118.I_LIST_ONHAND_STK_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										import_i_onhand_stock_tracking_no, _
										import_b_item, _
										import_b_storage_location_sl_cd, _
										import_b_plant_cd, _
										export_next_i_onhand_stock, _
										export_next_b_item_cd, _
										export_group, _
										Import_flag)
                                    
 If CheckSYSTEMError(Err, True) = True Then
	Set i22118 = Nothing
	Response.End
 End If

 Set i22118 = Nothing

 if IsEmpty(export_group) then
	Response.End
 End If
 
 LngMaxRow = CLng(Request("txtMaxRows"))
 ReDim PvArr(ubound(export_group,1))
 
 For LngRow = 0 To ubound(export_group,1)
     strData = Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_b_item_item_cd)) & _                                             
			   Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_b_item_item_nm)) & _                                              
			   Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_b_item_basic_unit)) & _                                           
			   Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_b_item_spec)) & _                                                 
			   Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_b_item_by_plant_location)) & _                                    
			   Chr(11) & ConvSPChars(export_group(LngRow,C_export_group_export_i_onhand_stock_tracking_no)) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_good_on_hand_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_bad_on_hand_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_stk_on_insp_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_stk_in_trns_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_schd_rcpt_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_schd_issue_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_prev_good_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_prev_bad_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_prev_stk_on_insp_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_prev_stk_in_trns_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_allocation_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 
			   Chr(11) & UniConvNumberDBToCompany(export_group(LngRow, C_export_group_export_i_onhand_stock_detail_picking_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _ 	 
			   Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
	PvArr(LngRow) = strData
 Next
 strData = Join(PvArr, "")

 If export_group(ubound(export_group,1), C_export_group_export_b_item_item_cd) = export_next_b_item_cd and _
    export_group(ubound(export_group,1), C_export_group_export_i_onhand_stock_tracking_no) = export_next_i_onhand_stock Then  

	StrNextKey1 = ""
	StrNextKey2 = ""
 Else
	StrNextKey1 = export_next_b_item_cd
	StrNextKey2 = export_next_i_onhand_stock
 End if


    Response.Write "<Script Language=vbscript>  " & vbCr   
    Response.Write " With Parent "                & vbCr
    
    Response.Write "   .ggoSpread.Source     = .frm1.vspdData "    & vbCr
    Response.Write "   .ggoSpread.SSShowData """ & strData  & """" & vbCr
    
    Response.Write "   .lgStrPrevKey1             = """ & ConvSPChars(StrNextKey1)    & """" & vbCr  
    Response.Write "   .lgStrPrevKey2             = """ & ConvSPChars(StrNextKey2)    & """" & vbCr  
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey1 <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
    
    Response.Write "End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr   
    
    Response.End 

%>
