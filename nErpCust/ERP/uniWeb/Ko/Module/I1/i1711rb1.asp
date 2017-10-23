<!--'**********************************************************************************************
'*  1. Module Name          : I Goods Movement detail List
'*  2. Function Name        : 
'*  3. Program ID           : I1711rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 수불 정보 상세 조회 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/05/14
'*  8. Modified date(Last)  : 2001/05/14
'*  9. Modifier (First)     : Hae Ryong Lee
'* 10. Modifier (Last)      : Hae Ryong Lee
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             
Err.Clear
On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE", "RB")

Call HideStatusWnd 

Dim pI11119
Dim strData
Dim PvArr  

Dim StrNextKey
Dim StrNextSubKey

Dim LngMaxRow
Dim LngRow
Dim GroupCount 

Const C_SHEETMAXROWS_D = 100


Dim I1_i_goods_movement_header
	Const I123_I1_item_document_no	= 0
	Const I123_I1_document_year		= 1
	Const I123_I1_trns_type			= 2
	Const I123_I1_plant_cd			= 3
ReDim I1_i_goods_movement_header(I123_I1_plant_cd)

Dim I2_i_goods_movement_detail
	Const I123_I2_seq_no		= 0
	Const I123_I2_sub_seq_no	= 1
ReDim I2_i_goods_movement_detail(I123_I2_sub_seq_no)


Dim E1_to_b_cost_center
Dim E2_b_cost_center_cost_nm
Dim E3_b_plant_plant_nm
Dim E4_b_minor_minor_nm
Dim E5_dest_b_item_item_nm
Dim E6_dest_b_plant_plant_nm
Dim E7_dest_b_storage_location_sl_nm
Dim E8_p_work_center_wc_nm

Dim E9_i_goods_movement_detail
	Const I123_E9_seq_no		= 0
	Const I123_E9_sub_seq_no	= 1 

Dim E10_i_goods_movement_header

Dim EG1_export_group
	Const I123_EG1_E1_b_storage_location_sl_nm				= 0
	Const I123_EG1_E2_b_item_item_cd						= 2
	Const I123_EG1_E2_b_item_item_nm						= 3
	Const I123_EG1_E3_i_goods_movement_detail_seq_no		= 5
	Const I123_EG1_E3_i_goods_movement_detail_sub_seq_no	= 6
	Const I123_EG1_E3_i_goods_movement_detail_qty			= 14
	Const I123_EG1_E3_i_goods_movement_detail_base_unit		= 15
	Const I123_EG1_E3_i_goods_movement_detail_price			= 16
	Const I123_EG1_E3_i_goods_movement_detail_amount		= 17
	Const I123_EG1_E3_i_goods_movement_detail_biz_area_cd	= 43

	StrNextKey    = Request("lgStrPrevKey")
	StrNextSubKey = Request("lgStrPrevSubKey")
 
	'-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_i_goods_movement_header(I123_I1_item_document_no)	= Request("txtItemDocumentNo")
    I1_i_goods_movement_header(I123_I1_document_year)		= Request("txtDocumentYear") 
	I1_i_goods_movement_header(I123_I1_trns_type)			= Request("txtTrnsType") 

    If StrNextKey <> "" AND StrNextSubKey <> "" Then
		I2_i_goods_movement_detail(I123_I2_seq_no)		= StrNextKey
		I2_i_goods_movement_detail(I123_I2_sub_seq_no)	= StrNextSubKey
	End If

	Set pI11119 = Server.CreateObject("PI0S140.cILookupGoodsMvmtSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End   
	End If    
 
	Call pI11119.CAB_I_LIST_GOODS_MVMT_DETAIL(gStrGlobalCollection, C_SHEETMAXROWS_D, _
											I1_i_goods_movement_header, _
											I2_i_goods_movement_detail, _
											E1_to_b_cost_center, _
											E2_b_cost_center_cost_nm, _
											E3_b_plant_plant_nm, _
											E4_b_minor_minor_nm, _
											E5_dest_b_item_item_nm, _
											E6_dest_b_plant_plant_nm, _
											E7_dest_b_storage_location_sl_nm, _
											E8_p_work_center_wc_nm, _
											E9_i_goods_movement_detail, _
											E10_i_goods_movement_header, _
											EG1_export_group)
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
		Set pI11119 = Nothing
		Response.End
	End If

	Set pI11119 = Nothing
 
	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows")) + 1

	if isEmpty(EG1_export_group) then
		Response.End
	End If
 
	GroupCount = ubound(EG1_export_group,1)
	ReDim PvArr(GroupCount)
	
	For LngRow = 0 To GroupCount
	    strData =	Chr(11) & ConvSPChars(EG1_export_group(LngRow, I123_EG1_E2_b_item_item_cd)) & _
		 			Chr(11) & ConvSPChars(EG1_export_group(LngRow, I123_EG1_E2_b_item_item_nm)) & _
		 			Chr(11) & ConvSPChars(EG1_export_group(LngRow, I123_EG1_E3_i_goods_movement_detail_base_unit)) & _
		 			Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I123_EG1_E3_i_goods_movement_detail_qty),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
		 			Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I123_EG1_E3_i_goods_movement_detail_price),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) & _
		 			Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I123_EG1_E3_i_goods_movement_detail_amount),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0) & _
		 			Chr(11) & ConvSPChars(EG1_export_group(LngRow, I123_EG1_E1_b_storage_location_sl_nm)) & _
		 			Chr(11) & ConvSPChars(EG1_export_group(LngRow, I123_EG1_E3_i_goods_movement_detail_biz_area_cd)) & _
		 			Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		PvArr(LngRow) = strData
	Next

	strData = Join(PvArr, "")
	
	If EG1_export_group(GroupCount, I123_EG1_E3_i_goods_movement_detail_seq_no) = E9_i_goods_movement_detail(I123_E9_seq_no) And _
		CInt(EG1_export_group(GroupCount, I123_EG1_E3_i_goods_movement_detail_sub_seq_no)) = CInt(E9_i_goods_movement_detail(I123_E9_sub_seq_no)) then
		
		StrNextKey    = ""
		StrNextSubKey = ""
	else       
		StrNextKey    = E9_i_goods_movement_detail(I123_E9_seq_no)
		StrNextSubKey = E9_i_goods_movement_detail(I123_E9_sub_seq_no)
	End If 

    Response.Write "<Script Language=vbscript> " & vbcr
    Response.Write "With parent "                & vbcr 
    Response.Write " .ggoSpread.Source = .vspdData "             & vbcr
    Response.Write " .ggoSpread.SSShowData """ & strData & """ " & vbcr
    Response.Write " .lgStrPrevKey =    """ & ConvSPChars(StrNextKey)    & """ " & vbcr
    Response.Write " .lgStrPrevSubKey = """ & ConvSPChars(StrNextSubKey) & """ " & vbcr

    Response.Write " if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKey <> """" Then "    & vbCr  
    Response.Write "    .DbQuery "                                                              & vbCr  
    Response.Write " else "                                                                     & vbCr  
    Response.Write "    .DbQueryOk "                                                            & vbCr  
    Response.Write " end if  "                                                                  & vbCr  

    Response.Write "End With "       & vbcr
    Response.Write "</Script> "      & vbcr
%>

