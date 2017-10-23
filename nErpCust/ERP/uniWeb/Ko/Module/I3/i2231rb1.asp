<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I2251rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 이월후 발생수불 
'*  6. Comproxy List        : 
'                             +I22511GoodsMovementCancelList
'*  7. Modified date(First) : 2001/11/08
'*  8. Modified date(Last)  : 2001/11/08
'*  9. Modifier (First)     : Sunggyu Han
'* 10. Modifier (Last)      : Sunggyu Han
'* 11. Comment              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             
on Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 

Err.Clear
Dim i22511            

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

Dim LngMaxRow
Dim LngRow
Dim isCount          
Dim strData
Dim PvArr

	Const C_SHEETMAXROWS_D = 100 

'-----------------------
'IMPORTS View
'-----------------------
Dim I1_b_plant
	Const I318_I1_plant_cd = 0
	Const I318_I1_inv_cls_dt = 1
ReDim I1_b_plant(I318_I1_inv_cls_dt)
Dim I2_i_goods_movement_cancel_list
	Const I318_I2_item_document_no = 0
	Const I318_I2_document_year = 1
	Const I318_I2_seq_no = 2
	Const I318_I2_sub_seq_no = 3
ReDim I2_i_goods_movement_cancel_list(I318_I2_sub_seq_no)
'-----------------------
'EXPORTS View
'-----------------------
Dim E1_ief_supplied_count
Dim E2_i_goods_movement_cancel_list
	Const I318_E2_item_document_no = 0
	Const I318_E2_document_year = 1
	Const I318_E2_seq_no = 2
	Const I318_E2_sub_seq_no = 3
Dim EG1_export_group
	Const I318_EG1_E1_item_document_no = 0
	Const I318_EG1_E1_document_year = 1
	Const I318_EG1_E1_seq_no = 2
	Const I318_EG1_E1_sub_seq_no = 3
	Const I318_EG1_E1_trns_type = 4
	Const I318_EG1_E1_mov_type = 5
	Const I318_EG1_E1_document_dt = 6
	Const I318_EG1_E1_pos_dt = 7
	Const I318_EG1_E1_post_flag = 8
	Const I318_EG1_E1_plant_cd = 9
	Const I318_EG1_E1_sl_cd = 10
	Const I318_EG1_E1_item_cd = 11
	Const I318_EG1_E1_lot_no = 12
	Const I318_EG1_E1_lot_sub_no = 13
	Const I318_EG1_E1_item_status = 14
	Const I318_EG1_E1_qty = 15
	Const I318_EG1_E1_price = 16
	Const I318_EG1_E1_amount = 17
	Const I318_EG1_E1_plan_order_no = 18
	Const I318_EG1_E1_req_no = 19
	Const I318_EG1_E1_prodt_order_no = 20
	Const I318_EG1_E1_dn_no = 21
	Const I318_EG1_E1_dn_seq = 22
	Const I318_EG1_E1_so_no = 23
	Const I318_EG1_E1_so_seq = 24
	Const I318_EG1_E1_po_no = 25
	Const I318_EG1_E1_po_seq_no = 26

	lgStrPrevKey1    = Request("lgStrPrevKeya1")
	lgStrPrevKey2    = Request("lgStrPrevKeya2")
	lgStrPrevKey3    = Request("lgStrPrevKeya3")
	lgStrPrevKey4    = Request("lgStrPrevKeya4")
	 
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------

	I1_b_plant(I318_I1_plant_cd) = Request("txtPlantCd")
	    
	I2_i_goods_movement_cancel_list(I318_I2_item_document_no) = lgStrPrevKey1
	I2_i_goods_movement_cancel_list(I318_I2_document_year)    = lgStrPrevKey2
	I2_i_goods_movement_cancel_list(I318_I2_seq_no)           = FilterVar(lgStrPrevKey3,0,"D")
	I2_i_goods_movement_cancel_list(I318_I2_sub_seq_no)       = FilterVar(lgStrPrevKey4,0,"D")

	Set i22511 = Server.CreateObject("PI3G070.cIGoodsMvtCancelList")
	    
	If CheckSYSTEMError(Err, True) = True Then
		Call ServerMesgBox("Object 1", vbCritical, I_MKSCRIPT)
		Response.End           
	End If    
	 
	Call i22511.I_GOODS_MVT_CANCEL_LIST(gStrGlobalCollection, C_SHEETMAXROWS_D, _
									I1_b_plant, _
									I2_i_goods_movement_cancel_list, _
									E1_ief_supplied_count, _
									E2_i_goods_movement_cancel_list, _
									EG1_export_group)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set i22511 = Nothing           
		Response.End              
	End If
	 
	Set i22511 = Nothing
	 
	If isEmpty(EG1_export_group) then
		Response.End              
	End If
	
	isCount	= UBOUND(EG1_export_group,1)
	strData = ""
	LngMaxRow = Request("txtMaxRows")
	Redim PvArr(UBOUND(EG1_export_group,1))

	For LngRow = 0 To isCount
		strData =	Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_item_document_no)) & _
					Chr(11) & UNIDateClientFormat(EG1_export_group(LngRow, I318_EG1_E1_document_dt))
	        
		Select Case  ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_trns_type))
		Case "PR"
			strData = strData & Chr(11) & "구매입고"
		Case "MR"
			strData = strData & Chr(11) & "생산입고"    
		Case "OR"
			strData = strData & Chr(11) & "예외입고"     
		Case "PI"
			strData = strData & Chr(11) & "생산출고"     
		Case "DI"
			strData = strData & Chr(11) & "판매출고"     
		Case "OI"
			strData = strData & Chr(11) & "예외출고"     
		Case "ST"
			strData = strData & Chr(11) & "재고이동"         
		End Select  
		        
		strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_sl_cd)) & _
							Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_item_cd)) & _
							Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_lot_no)) & _
							Chr(11) & EG1_export_group(LngRow, I318_EG1_E1_lot_sub_no) & _
							Chr(11) & EG1_export_group(LngRow, I318_EG1_E1_seq_no) & _
							Chr(11) & EG1_export_group(LngRow, I318_EG1_E1_sub_seq_no) & _
							Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_dn_no)) & _
							Chr(11) & ConvSPChars(EG1_export_group(LngRow, I318_EG1_E1_po_no)) & _
							Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I318_EG1_E1_qty),  ggQty.DecPoint,      ggQty.RndPolicy,      ggQty.RndUnit,0) & _
							Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I318_EG1_E1_price),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0) & _
							Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)        
							
		PvArr(LngRow) = strData
			
	Next
	
		strData = Join(PvArr, "")

	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If
	
	If  EG1_export_group(isCount, I318_EG1_E1_item_document_no)		= E2_i_goods_movement_cancel_list(I318_E2_item_document_no) _
		And EG1_export_group(isCount, I318_EG1_E1_document_year)    = E2_i_goods_movement_cancel_list(I318_E2_document_year) _
		And EG1_export_group(isCount, I318_EG1_E1_seq_no)           = E2_i_goods_movement_cancel_list(I318_E2_seq_no) _
		And EG1_export_group(isCount, I318_EG1_E1_sub_seq_no)       = E2_i_goods_movement_cancel_list(I318_E2_sub_seq_no) then      
	 
		lgStrPrevKey1 = ""
		lgStrPrevKey2 = ""
		lgStrPrevKey3 = ""
		lgStrPrevKey4 = "" 
	else
		lgStrPrevKey1 = ConvSPChars(E2_i_goods_movement_cancel_list(I318_E2_item_document_no))
		lgStrPrevKey2 = ConvSPChars(E2_i_goods_movement_cancel_list(I318_E2_document_year))
		lgStrPrevKey3 = ConvSPChars(E2_i_goods_movement_cancel_list(I318_E2_seq_no))
		lgStrPrevKey4 = ConvSPChars(E2_i_goods_movement_cancel_list(I318_E2_sub_seq_no))
	End If    


	Response.Write "<Script Language=vbscript> " & vbcr
	Response.Write "With parent "                & vbcr      '☜: 화면 처리 ASP 를 지칭함 
	 
	Response.Write " .ggoSpread.Source = .vspdData " & vbcr
	Response.Write " .ggoSpread.SSShowData """ & strData & """ " & vbcr

	Response.Write " .lgStrPrevKey1 = """ & lgStrPrevKey1 & """ " & vbcr
	Response.Write " .lgStrPrevKey2 = """ & lgStrPrevKey2 & """ " & vbcr
	Response.Write " .lgStrPrevKey3 = """ & lgStrPrevKey3 & """ " & vbcr
	Response.Write " .lgStrPrevKey4 = """ & lgStrPrevKey4 & """ " & vbcr
	Response.Write " If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey1 <> """" And .lgStrPrevKey2 <> """" And .lgStrPrevKey3 <> """" And .lgStrPrevKey4 <> """" Then " & vbcr
	Response.Write "  .DbQuery "                                & vbcr
	Response.Write " Else "                                     & vbcr
	Response.Write "  .DbQueryOk "								& vbcr
	Response.Write " End If "                                   & vbcr

	Response.Write "End with " & vbcr
	Response.Write "</Script> " & vbcr
%> 
 
