<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Maint Phy inv (Manual)
'*  3. Program ID           : I2121mb1.asp
'*  4. Program Name         : 실사선별조정 
'*  5. Program Desc         : 선별된 품목에 대하여 조회한다.
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2003/06/02
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
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
Dim pPI2S030												

Dim strData
Dim StrNextKey1	
Dim StrNextKey2	

Dim LngMaxRow	
Dim LngRow    
Dim PvArr

Const C_SHEETMAXROWS_D = 100
    '-----------------------
    'IMPORTS View
    '-----------------------
    Dim I1_b_plant_cd
    Dim I2_i_physical_inventory_header_phy_inv_no
    Dim I3_i_physical_inventory_detail_seq_no
    Dim I4_b_item_cd
    Dim I5_display_flg 
    
	'-----------------------
	'EXPORTS View
	'-----------------------
    Dim EG1_group_export
		Const I209_EG1_E1_good_mvmt_workset_amount = 0
		Const I209_EG1_E1_good_mvmt_workset_entry_qty = 1
		Const I209_EG1_E2_i_onhand_stock_tracking_no = 2
		Const I209_EG1_E3_i_onhand_stock_detail_lot_no = 3
		Const I209_EG1_E3_i_onhand_stock_detail_lot_sub_no = 4
		Const I209_EG1_E4_b_item_item_cd = 5
		Const I209_EG1_E4_b_item_item_nm = 6
		Const I209_EG1_E4_b_item_spec = 7
		Const I209_EG1_E4_b_item_basic_unit = 8
		Const I209_EG1_E5_i_physical_inventory_detail_seq_no = 9
		Const I209_EG1_E5_i_physical_inventory_detail_real_insp_adj_dt = 10
		Const I209_EG1_E5_i_physical_inventory_detail_prc = 11
		Const I209_EG1_E5_i_physical_inventory_detail_abc_flag = 12
		Const I209_EG1_E5_i_physical_inventory_detail_spcl_stk_indctr = 13
		Const I209_EG1_E5_i_physical_inventory_detail_sts_indctr = 14
		Const I209_EG1_E5_i_physical_inventory_detail_bad_qty = 15
		Const I209_EG1_E5_i_physical_inventory_detail_good_qty = 16
		Const I209_EG1_E5_i_physical_inventory_detail_inv_good_qty = 17
		Const I209_EG1_E5_i_physical_inventory_detail_inv_bad_qty = 18
		Const I209_EG1_E5_i_physical_inventory_detail_zero_cnt_indctr = 19
		Const I209_EG1_E5_i_physical_inventory_detail_cycle_cnting_indctr = 20    

    Dim E2_i_physical_inventory_detail_seq_no
    Dim E3_b_item_item_cd


    StrNextKey1 = Request("lgStrPrevKey1")
    StrNextKey2 = Request("lgStrPrevKey2")
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_b_plant_cd                             = Request("txtPlantCd")
    I2_i_physical_inventory_header_phy_inv_no = Request("txtCondPhyInvNo")
    I3_i_physical_inventory_detail_seq_no     = ""
    I4_b_item_cd                              = Request("txtItemCd")    
    I5_display_flg							  = "ML"	
    
    if StrNextKey1 <> "" and StrNextKey2 <> "" then
		I3_i_physical_inventory_detail_seq_no   = StrNextKey1
    	I4_b_item_cd							= StrNextKey2
    end if

	
	Set pPI2S030 = Server.CreateObject("PI2S030.cILookupPhyInv")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    
	
	Call pPI2S030.I_LOOKUP_PHY_INV(gStrGlobalCollection, C_SHEETMAXROWS_D, _
									I1_b_plant_cd, _
									I2_i_physical_inventory_header_phy_inv_no, _
									I3_i_physical_inventory_detail_seq_no, _
									I4_b_item_cd, _
									I5_display_flg, _
									EG1_group_export, _
									E2_i_physical_inventory_detail_seq_no, _
									E3_b_item_item_cd)

    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI2S030 = Nothing										
		Response.End											
	End If

	
	Set pPI2S030 = Nothing
	
	if isEmpty(EG1_group_export) then
		Response.End													
	end if
	
	LngMaxRow = CLng(Request("txtMaxRows")) + 1

	ReDim PvArr(ubound(EG1_group_export,1))

	StrNextKey1 = E2_i_physical_inventory_detail_seq_no
	StrNextKey2 = E3_b_item_item_cd

	For LngRow = 0 To ubound(EG1_group_export,1)
	
		strData = Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E4_b_item_item_cd)) & _
				  Chr(11) & "" & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow ,I209_EG1_E4_b_item_item_nm)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E4_b_item_spec)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E4_b_item_basic_unit)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E2_i_onhand_stock_tracking_no)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E3_i_onhand_stock_detail_lot_no)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E3_i_onhand_stock_detail_lot_sub_no)) & _
				  Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E5_i_physical_inventory_detail_abc_flag))
		If UCase(Trim(EG1_group_export(LngRow, I209_EG1_E5_i_physical_inventory_detail_sts_indctr))) = "C" Then
			strData = strData & Chr(11) & "Y"
		Else 
			strData = strData & Chr(11) & "N"
		End If				
		
		strData = strData & Chr(11) & ConvSPChars(EG1_group_export(LngRow, I209_EG1_E5_i_physical_inventory_detail_seq_no)) & _
							Chr(11) & LngMaxRow + LngRow & Chr(11)

        PvArr(LngRow) = strData

	Next
	
    strData = Join(PvArr, Chr(12)) & Chr(12)

    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr

	Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"        & vbCr
	
	Response.Write "    .lgStrPrevKey1  = """ & ConvSPChars(StrNextKey1) & """" & vbCr  
    Response.Write "    .lgStrPrevKey2  = """ & ConvSPChars(StrNextKey2) & """" & vbCr  
	
	Response.Write "	.frm1.txthCondPhyInvNo.value = """ & ConvSPChars(Request("txtCondPhyInvNo")) & """" & vbCr  	   	  
  	Response.Write "	.frm1.txthPlantCd.value      = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr

    Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey1 <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	
	Response.Write "End with	" & vbCr
    Response.Write "</Script>      " & vbCr   

	Response.End     
%>
