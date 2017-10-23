<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>

<!--'********************************************************************************************************
'*  1. Module Name          : Inventory               *
'*  2. Function Name        : 수불상세         *
'*  3. Program ID           : i2241rb1.asp                *
'*  4. Program Name         :                   *
'*  5. Program Desc         : 수불상세팝업                 *
'*  7. Modified date(First) : 2000/05/02                *
'*  8. Modified date(Last)  : 2000/05/02                *
'*  9. Modifier (First)     :                 *
'* 10. Modifier (Last)      :                 *
'* 11. Comment              :                   *
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"         *
'*                            this mark(⊙) Means that "may  change"         *
'*                            this mark(☆) Means that "must change"         *
'* 13. History              :                   *
'*                            2000/05/02 : Coding Start             *
'********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                   
err.Clear
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")  
Call HideStatusWnd 

On Error Resume Next

Dim objPopUp

Dim StrData
Dim strMode                
Dim intColCnt              
Dim LngRow
Dim LngMaxRows
Dim intGroupCount 
Dim StrNextKey
Dim PvArr 

	Const C_SHEETMAXROWS_D = 100

'-----------------------
'IMPORTS View
'-----------------------
Dim I1_b_plant_cd
Dim I2_i_monthly_inventory
	Const I317_I2_mnth_inv_year = 0
	Const I317_I2_mnth_inv_month = 1
ReDim I2_i_monthly_inventory(I317_I2_mnth_inv_month)
Dim I3_fr_b_item
	Const I317_I3_item_cd = 0
	Const I317_I3_item_acct = 1
	Const I317_I3_item_class = 2
ReDim I3_fr_b_item(I317_I3_item_class)
Dim I4_to_b_item_cd
Dim I5_next_b_item_cd
Dim I6_next_i_inventory_history_for_mov_type
'-----------------------
'EXPORTS View
'-----------------------
Dim E1_b_plant
Dim EG1_export_group
	Const I317_EG1_E1_item_cd = 0
	Const I317_EG1_E1_item_nm = 1
	Const I317_EG1_E1_spec = 2
	Const I317_EG1_E1_basic_unit = 3
	Const I317_EG1_E1_item_acct = 4
	Const I317_EG1_E1_item_class = 5
	Const I317_EG1_E2_mov_type = 6
	Const I317_EG1_E2_qty = 7
	Const I317_EG1_E2_amount = 8
	Const I317_EG1_E3_minor_nm = 9
Dim E2_next_b_item_cd
Dim E3_next_i_inventory_history_for_mov_type


	StrNextKey       = Request("lgStrPrevKey")
	LngMaxRows       = Request("txtMaxRows")
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	I1_b_plant_cd									= Request("txtPlantCd")
	I3_fr_b_item(I317_I3_item_cd)					= Request("txtItemCd")
	I3_fr_b_item(I317_I3_item_acct)					= ""
	I3_fr_b_item(I317_I3_item_class)				= ""
	I4_to_b_item_cd									= Request("txtItemCd")
	I2_i_monthly_inventory(I317_I2_mnth_inv_year)	= Request("txtYyyy")
	I2_i_monthly_inventory(I317_I2_mnth_inv_month)  = Request("txtMm")

	If StrNextKey <> "" Then I6_next_i_inventory_history_for_mov_type = StrNextKey
	                                 
	Set objPopUp = Server.CreateObject("PI3G060.cIListInvHistSvr")
	    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End            
	End If    
	 
	Call objPopUp.I_LIST_INV_HIST_FOR_MOVETYPE(gStrGlobalCollection, C_SHEETMAXROWS_D, _
												I1_b_plant_cd, _
												I2_i_monthly_inventory, _
												I3_fr_b_item, _
												I4_to_b_item_cd, _
												I5_next_b_item_cd, _
												I6_next_i_inventory_history_for_mov_type, _
												E1_b_plant, _
												EG1_export_group, _
												E2_next_b_item_cd, _
												E3_next_i_inventory_history_for_mov_type)
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set objPopUp = Nothing           
		Response.End              
	End If

	Set objPopUp = Nothing
	 
	If isEmpty(EG1_export_group) then
		Response.End              
	End If


	intGroupCount = ubound(EG1_export_group,1)
	
	strData = ""
	Redim PvArr(ubound(EG1_export_group,1))

	For LngRow = 0 To intGroupCount
		strData =	Chr(11) & ConvSPChars(EG1_export_group(LngRow, I317_EG1_E2_mov_type)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I317_EG1_E3_minor_nm)) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I317_EG1_E2_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I317_EG1_E2_amount),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0) & _
					Chr(11) & LngMaxRows + LngRow & Chr(11) & Chr(12) 
		PvArr(LngRow) = strData
	Next
		strData = Join(PvArr, "")

	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If

	If EG1_export_group(intGroupCount, I317_EG1_E2_mov_type) = E3_next_i_inventory_history_for_mov_type then
		StrNextKey = ""
	Else
		StrNextKey = E3_next_i_inventory_history_for_mov_type 
	End If


	Response.Write "<Script Language=vbscript> "                & vbcr
	Response.Write "With parent "                               & vbcr
	Response.Write " .ggoSpread.Source = .vspdData "            & vbcr 
	Response.Write " .ggoSpread.SSShowData """ & strData & """" & vbcr
	    
	Response.Write " .lgStrPrevKey = """ & ConvSPChars(StrNextKey) & """"                    & vbcr
	Response.Write " If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbcr
	Response.Write "  .DbQuery "                                                             & vbcr
	Response.Write " Else "                                                                  & vbcr
	Response.Write "  .DbQueryOk "                                                           & vbcr
	Response.Write " End If "                                                                & vbcr
	    
	Response.Write "End With " & vbcr
	Response.Write "</Script> " & vbcr
%>
