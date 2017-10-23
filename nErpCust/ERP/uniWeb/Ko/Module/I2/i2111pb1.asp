<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Phy Inv header (Manual)
'*  3. Program ID           : I2111pb1.asp
'*  4. Program Name         : 실사번호조회 
'*  5. Program Desc         : 실사번호정보를  조회한다.
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
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
Call LoadBasisGlobalInf()

Err.Clear
On Error Resume Next												
Call HideStatusWnd 
Dim pPI2G040													
Dim PvArr
Dim strData													

Dim StrNextKey		
Dim LngMaxRow		
Dim LngRow
	
Const C_SHEETMAXROWS_D = 100
Dim lgStrPrevKey

    Dim I1_b_plant_cd
    Dim I2_b_storage_location_sl_cd
    Dim I3_i_physical_inventory_header
		Const I212_I3_phy_inv_no = 0
		Const I212_I3_doc_sts_indctr = 1    
    ReDim I3_i_physical_inventory_header(I212_I3_doc_sts_indctr)
    Dim I4_to_prod_work_set_temp_timestamp
    Dim I5_fr_prod_work_set_temp_timestamp
    
    Dim E1_i_physical_inventory_header_phy_inv_no
    Dim EG1_export_group
		Const I212_EG1_E1_b_plant_plant_cd = 0
		Const I212_EG1_E1_b_plant_plant_nm = 1
		Const I212_EG1_E2_b_storage_location_sl_cd = 2
		Const I212_EG1_E2_b_storage_location_sl_nm = 3
		Const I212_EG1_E3_i_physical_inventory_header_phy_inv_no = 4
		Const I212_EG1_E3_i_physical_inventory_header_real_insp_dt = 5
		Const I212_EG1_E3_i_physical_inventory_header_pos_blk_indctr = 6

    StrNextKey = Request("lgStrPrevKey")
	
	I3_i_physical_inventory_header(I212_I3_phy_inv_no)     = Request("txtPhyInvNo")
	I3_i_physical_inventory_header(I212_I3_doc_sts_indctr) = Request("lgDocSts")
	I2_b_storage_location_sl_cd                            = Request("txtSLCd")
	I1_b_plant_cd                                          = Request("txtPlantCd")
	I5_fr_prod_work_set_temp_timestamp                     = UniConvDate(Request("txtFromDt"))
	
    If Trim(Request("txtToDt")) = "" then
		I4_to_prod_work_set_temp_timestamp = ""
	else
		I4_to_prod_work_set_temp_timestamp = UniConvDate(Request("txtToDt"))
	End If

    if StrNextKey <> "" then I3_i_physical_inventory_header(I212_I3_phy_inv_no) = StrNextKey

	Set pPI2G040 = Server.CreateObject("PI2G040.cIListPhyInvHdr")

	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    
	
	Call pPI2G040.I_LIST_PHY_INV_HEADER(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_b_plant_cd, _
										I2_b_storage_location_sl_cd, _
										I3_i_physical_inventory_header, _
										I4_to_prod_work_set_temp_timestamp, _
										I5_fr_prod_work_set_temp_timestamp, _
										E1_i_physical_inventory_header_phy_inv_no, _
										EG1_export_group)

    If CheckSYSTEMError(Err, True) = True Then
    	Set pPI2G040 = Nothing											
		Response.End													
	End If

	Set pPI2G040 = Nothing

	if isEmpty(EG1_export_group) then
		Response.End													
	end if
	
	ReDim PvArr(ubound(EG1_export_group,1))
	LngMaxRow = CLng(Request("txtMaxRows")) + 1

	For LngRow = 0 To ubound(EG1_export_group,1)
		PvArr(LngRow) = Chr(11) & ConvSPChars(EG1_export_group(LngRow, I212_EG1_E3_i_physical_inventory_header_phy_inv_no)) & _
						Chr(11) & UNIDateClientFormat(EG1_export_group(LngRow, I212_EG1_E3_i_physical_inventory_header_real_insp_dt)) & _
						Chr(11) & ConvSPChars(EG1_export_group(LngRow, I212_EG1_E2_b_storage_location_sl_cd)) & _
						Chr(11) & ConvSPChars(EG1_export_group(LngRow, I212_EG1_E2_b_storage_location_sl_nm)) & _
						Chr(11) & EG1_export_group(LngRow, I212_EG1_E3_i_physical_inventory_header_pos_blk_indctr) & _
						Chr(11) & ConvSPChars(EG1_export_group(LngRow, I212_EG1_E1_b_plant_plant_cd)) & _
						Chr(11) & ConvSPChars(EG1_export_group(LngRow, I212_EG1_E1_b_plant_plant_nm)) & _
						Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
	Next
    
    strData = Join(PvArr, "")
	
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If

	If EG1_export_group(ubound(EG1_export_group,1), I212_EG1_E3_i_physical_inventory_header_phy_inv_no) = E1_i_physical_inventory_header_phy_inv_no Then	 
		StrNextKey = ""
	Else
		StrNextKey = E1_i_physical_inventory_header_phy_inv_no
	End if
    
    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr

    Response.Write "   .ggoSpread.Source     = .vspdData "                          & vbCr
    Response.Write "   .ggoSpread.SSShowData """ & strData  & """"                  & vbCr
    Response.Write "   .vspdData.focus "                                            & vbCr
    Response.Write "   .hlgDocumentNo = """ & ConvSPChars(Request("txtPhyInvNo")) & """" & vbCr  
    Response.Write "   .hlgDocSts = """ & ConvSPChars(Request("lgDocSts")) & """" & vbCr  
    Response.Write "   .hlgSLCd = """ & ConvSPChars(Request("txtSLCd")) & """" & vbCr  
    Response.Write "   .hlgPlantCd = """ &ConvSPChars(Request("txtPlantCd")) & """" & vbCr  
    Response.Write "   .hlgFromDt = """ & Request("txtFromDt") & """" & vbCr  
    Response.Write "   .hlgToDt = """ & Request("txtToDt") & """" & vbCr  

    Response.Write "   .lgStrPrevKey  = """ & ConvSPChars(StrNextKey) & """" & vbCr  
    Response.Write " if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKey <> """" Then "    & vbCr  
    Response.Write "    .DbQuery "                                                              & vbCr  
    Response.Write " else "                                                                     & vbCr  
    Response.Write "    .DbQueryOk "                                                            & vbCr  
    Response.Write " end if  "                                                                  & vbCr  
    
    Response.Write "End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr   
    
    Response.End 
%>

