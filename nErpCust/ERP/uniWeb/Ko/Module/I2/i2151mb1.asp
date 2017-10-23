<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Create physical inventory Posting in batch 
'*  3. Program ID           : I2151mb1.asp
'*  4. Program Name         : 실사조정Batch등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2000/04/06
'*  8. Modified date(Last)  : 2002/07/06
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%									
Call LoadBasisGlobalInf()

On Error Resume Next
Err.Clear
Call HideStatusWnd

Dim pPI2G070												

Dim strMode

Dim I1_b_storage_location_sl_cd 
Dim I2_b_plant_plant_cd 
Dim I3_i_physical_inventory_header_phy_inv_no
Dim I4_b_cost_center_cost_cd

Dim E1_good_mvmt_workset
	Const E1_item_document_no = 0   
	Const E1_year = 1

Dim prErrorPosition
	Const Err_item_cd = 0
    Const Err_tracking_no = 1
    Const Err_lot_no = 2
    Const Err_lot_sub_no = 3
 
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_b_plant_plant_cd					= Request("txtPlantCd")
    I1_b_storage_location_sl_cd			= Request("txtSLCd")
    I4_b_cost_center_cost_cd            = Request("txtCostCd")        
    I3_i_physical_inventory_header_phy_inv_no       = Request("txtPhyinvNo")
 
 
    Set pPI2G070 = Server.CreateObject("PI2G070.cIPostPhyInvBatch")    

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

    '-----------------------
    'Com action area
    '-----------------------
	Call pPI2G070.I_POST_PHY_INV_BATCH(gStrGlobalCollection, _
										I1_b_storage_location_sl_cd, _
										I2_b_plant_plant_cd, _
										I3_i_physical_inventory_header_phy_inv_no, _
										I4_b_cost_center_cost_cd, _
										E1_good_mvmt_workset, _
										prErrorPosition)
    
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		If prErrorPosition(0) <> "" Then
			Call ServerMesgBox("상세정보" & vbcrlf & vbcrlf & vbcrlf & _
							   "품목 : " & prErrorPosition(Err_item_cd) & vbtab & vbtab & vbcrlf & vbcrlf & _
							   "Tracking No : " & prErrorPosition(Err_tracking_no) & vbtab & vbcrlf &vbcrlf & _
							   "LOT NO : " & prErrorPosition(Err_lot_no) & vbcrlf & vbtab & vbcrlf & _
							   "Lot No.순번 : " & prErrorPosition(Err_lot_sub_no), _
							   	vbCritical, I_MKSCRIPT)  
		End If

		Set pPI2G070 = Nothing														
		Response.End
	End If

    Set pPI2G070 = Nothing	

	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.hItemDocumentNo.value = """ & ConvSPChars(E1_good_mvmt_workset(E1_item_document_no)) & """" & vbCr  	   	  
  	Response.Write "    .DbSaveOk "				& vbCr
	Response.Write " End with	" & vbCr
    Response.Write "</Script>      " & vbCr   
	Response.End 
%>