<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Good Movement Header/detail
'*  3. Program ID           : I1131mb2.asp
'*  4. Program Name         : 기타입고수불등록 
'*  5. Program Desc         : 기타입고수불정보/상세정보를 등록한다.
'*  7. Modified date(First) : 2002/05/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HAN SUNG GYU
'* 10. Modifier (Last)      : HAN SUNG GYU
'* 11. Comment              : VB CONVERSION시 반영 
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
On Error Resume Next													

Call LoadBasisGlobalInf()
    
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Call HideStatusWnd 

Dim iPI0S140											
Dim strMode	
											
Dim StrNextKey		
Dim lgStrPrevKey	
Dim LngMaxRow		
Dim LngRow
Dim GroupCount          
Const MovType = "OR"		

Dim I1_i_good_mvmt_header 
Const C_I1_ItemDocumentNo	= 0 
Const C_I1_DocumentYear		= 1
Const C_I1_TrnsType			= 2
Const C_I1_PlantCd			= 3

Redim  I1_i_good_mvmt_header(C_I1_PlantCd)

Dim I2_i_good_mvmt_detail
Const C_I2_SeqNo	= 0 
Const C_I2_SubSeqNo = 1

Redim I2_i_good_mvmt_detail(C_I2_SubSeqNo)

Const C_SHEETMAXROWS_D  =   100

Dim strData
Dim E1_b_cost_center,    E2_b_cost_center_cost_nm , _
    E3_b_plant_plant_nm ,  E4_b_minor_minor_nm , _
    E5_b_item_item_nm , E6_b_plant_plant_nm , _
    E7_b_storage_location_sl_nm , E8_p_work_center_wc_nm , _
    E9_i_goods_movement_detail ,      E10_i_goods_movement_header ,  _
    EG1_export_group                
    
Const C_E1_cost_cd = 0
Const C_E1_cost_nm = 1

Const C_E9_seq_no		= 0
Const C_E9_sub_seq_no	= 1
    
Const C_E10_item_document_no	= 0
Const C_E10_document_year		= 1
Const C_E10_trns_type			= 2
Const C_E10_mov_type			= 3
Const C_E10_document_dt			= 4
Const C_E10_pos_dt				= 5
Const C_E10_document_text		= 6
Const C_E10_plant_cd			= 7
Const C_E10_cost_cd				= 8
   
Const C_EG1_b_storage_location_sl_nm							= 0
Const C_EG1_b_storage_location_sl_cd							= 1
Const C_EG1_b_item_item_cd										= 2
Const C_EG1_b_item_item_nm										= 3
Const C_EG1_b_item_spec											= 4
Const C_EG1_i_goods_movement_detail_seq_no						= 5
Const C_EG1_i_goods_movement_detail_sub_seq_no					= 6
Const C_EG1_i_goods_movement_detail_auto_crtd_flag				= 7
Const C_EG1_i_goods_movement_detail_debit_credit_flag			= 8
Const C_EG1_i_goods_movement_detail_trns_type					= 9
Const C_EG1_i_goods_movement_detail_mov_type					= 10
Const C_EG1_i_goods_movement_detail_lot_no						= 11
Const C_EG1_i_goods_movement_detail_lot_sub_no					= 12
Const C_EG1_i_goods_movement_detail_item_status					= 13
Const C_EG1_i_goods_movement_detail_qty							= 14
Const C_EG1_i_goods_movement_detail_base_unit					= 15
Const C_EG1_i_goods_movement_detail_price						= 16
Const C_EG1_i_goods_movement_detail_amount						= 17
Const C_EG1_i_goods_movement_detail_cost_of_devy				= 18
Const C_EG1_i_goods_movement_detail_cur_cd						= 19
Const C_EG1_i_goods_movement_detail_entry_qty					= 20
Const C_EG1_i_goods_movement_detail_entry_unit					= 21
Const C_EG1_i_goods_movement_detail_order_qty					= 22
Const C_EG1_i_goods_movement_detail_order_unit					= 23
Const C_EG1_i_goods_movement_detail_stock_type					= 24
Const C_EG1_i_goods_movement_detail_bp_cd						= 25
Const C_EG1_i_goods_movement_detail_wc_cd						= 26
Const C_EG1_i_goods_movement_detail_trns_lot_no					= 27
Const C_EG1_i_goods_movement_detail_trns_lot_sub_no				= 28
Const C_EG1_i_goods_movement_detail_trns_plant_cd				= 29
Const C_EG1_i_goods_movement_detail_trns_sl_cd					= 30
Const C_EG1_i_goods_movement_detail_trns_item_cd				= 31
Const C_EG1_i_goods_movement_detail_plan_order_no				= 32
Const C_EG1_i_goods_movement_detail_req_no						= 33
Const C_EG1_i_goods_movement_detail_prodt_order_no				= 34
Const C_EG1_i_goods_movement_detail_dn_no						= 35
Const C_EG1_i_goods_movement_detail_dn_seq						= 36
Const C_EG1_i_goods_movement_detail_so_no						= 37
Const C_EG1_i_goods_movement_detail_so_seq						= 38
Const C_EG1_i_goods_movement_detail_po_no						= 39
Const C_EG1_i_goods_movement_detail_po_seq_no					= 40
Const C_EG1_i_goods_movement_detail_tracking_no					= 41
Const C_EG1_i_goods_movement_detail_trns_tracking_no			= 42
Const C_EG1_i_goods_movement_detail_biz_area_cd					= 43
Const C_EG1_i_goods_movement_detail_branch_flag					= 44
Const C_EG1_i_goods_movement_detail_pur_order_type				= 45
Const C_EG1_i_goods_movement_detail_subcntrct_mfg_cost_amount	= 46
Const C_EG1_i_goods_movement_detail_cost_cd						= 47

Dim PvArr

    if Request("lgStrPrevKey") = "" then
     lgStrPrevKey = 0
    Else
     lgStrPrevKey = Request("lgStrPrevKey")
    End if
    
    I1_i_good_mvmt_header(C_I1_ItemDocumentNo)	= Trim(Request("txtDocumentNo1"))
    I1_i_good_mvmt_header(C_I1_DocumentYear)	= Trim(Request("txtYear"))
    I1_i_good_mvmt_header(C_I1_TrnsType)		= "OR"    
    I1_i_good_mvmt_header(C_I1_PlantCd)			= Trim(Request("txtPlantCd"))
    
    If isnull(lgStrPrevKey) then
       I2_i_good_mvmt_detail(C_I2_SeqNo) = 0 
    else 
       I2_i_good_mvmt_detail(C_I2_SeqNo) = lgStrPrevKey
    End If 
    
    Set iPI0S140 = Server.CreateObject("PI0S140.cILookupGoodsMvmtSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
          Set PIG1010 = Nothing											
		  Response.End													
    End If
    
    call iPI0S140.CAB_I_LIST_GOODS_MVMT_DETAIL(gStrGlobalCollection,C_SHEETMAXROWS_D,_
											I1_i_good_mvmt_header, _
											I2_i_good_mvmt_detail, _
											E1_b_cost_center, _
											E2_b_cost_center_cost_nm , _
											E3_b_plant_plant_nm, _
											E4_b_minor_minor_nm, _
											E5_b_item_item_nm, _
											E6_b_plant_plant_nm , _
											E7_b_storage_location_sl_nm, _
											E8_p_work_center_wc_nm , _
											E9_i_goods_movement_detail, _
											E10_i_goods_movement_header, _
											EG1_export_group)
   
    If CheckSYSTEMError(Err,True) = True Then
		Set PI1G010 = Nothing											
		Response.End													
    End If

    Set iPI0S140 = Nothing

	ReDim PvArr(UBound(EG1_export_group,1))
	
    For LngRow = 0 To UBound(EG1_export_group,1)
		strData = Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_b_item_item_cd)) & _
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_b_item_item_nm)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_amount),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _ 
				  Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_entry_qty),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_entry_unit)) & _ 
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_base_unit)) & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_b_item_spec)) & _ 
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_tracking_no)) & _ 
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_lot_no)) & _ 
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_lot_sub_no)) & _ 
				  Chr(11) & "" & _ 
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_seq_no)) & _ 
				  Chr(11) & ConvSPChars(EG1_export_group(LngRow,C_EG1_i_goods_movement_detail_sub_seq_no)) & _ 
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
    
		PvArr(LngRow) = strData
    Next

	strData = Join(PvArr, "")
	If EG1_export_group(ubound(EG1_export_group,1), C_EG1_i_goods_movement_detail_seq_no) = E9_i_goods_movement_detail(C_E9_seq_no) Then	 
		StrNextKey = ""
	Else
		StrNextKey = E9_i_goods_movement_detail(C_E9_seq_no)
	End if

    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent" & vbCr
	Response.Write "	.frm1.txtDocumentDt.text	= """ & UNIDateClientFormat(E10_i_goods_movement_header(C_E10_document_dt)) & """" & vbCr
	Response.Write "	.frm1.txtPostingDt.text     = """ & UNIDateClientFormat(E10_i_goods_movement_header(C_E10_pos_dt)) & """" & vbCr
	Response.Write "	.frm1.txtMovType.value      = """ & ConvSPChars(E10_i_goods_movement_header(C_E10_mov_type)) & """" & vbCr
	Response.Write "	.frm1.txtMovTypeNm.value    = """ & ConvSPChars(E4_b_minor_minor_nm(0)) & """" & vbCr
	Response.Write "	.frm1.txtSLCd.value         = """ & ConvSPChars(EG1_export_group(0,C_EG1_b_storage_location_sl_cd)) & """" & vbCr
	Response.Write "	.frm1.txtSLNm.value         = """ & ConvSPChars(EG1_export_group(0,C_EG1_b_storage_location_sl_nm)) & """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value      = """ & ConvSPChars(E3_b_plant_plant_nm(0)) & """" & vbCr
	Response.Write "	.frm1.txtDocumentText.value = """ & ConvSPChars(E10_i_goods_movement_header(C_E10_document_text)) & """" & vbCr
	Response.Write "	.frm1.txtDocumentNo2.value  = """ & ConvSPChars(E10_i_goods_movement_header(C_E10_item_document_no)) & """" & vbCr	
	Response.Write "	.frm1.txtDocumentNo1.value  = """ & ConvSPChars(E10_i_goods_movement_header(C_E10_item_document_no)) & """" & vbCr	
	Response.Write "	.frm1.txtCostCd.value       = """ & ConvSPChars(E10_i_goods_movement_header(C_E10_cost_cd)) & """" & vbCr
	Response.Write "	.frm1.txtCostNm.value       = """ & ConvSPChars(E2_b_cost_center_cost_nm(0)) & """" & vbCr
	Response.Write "	.lgStrPrevKey				= """ & ConvSPChars(StrNextKey)    & """" & vbCr

	Response.Write "	.ggoSpread.Source = .frm1.vspdData		  " & vbcr
	Response.Write "	.ggoSpread.SSShowData """ & strData & """ " & vbcr
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
    Response.Write "End with " & vbcr
	Response.Write "</Script>	" & vbCr
	
	Response.End													
%>
