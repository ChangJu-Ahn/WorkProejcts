<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Inventory Analsys
'*  3. Program ID           : I2213mb1.asp
'*  4. Program Name         : List Stock Requirement
'*  5. Program Desc         : 품목재고정보 자료를 조회한다.
'*  6. Comproxy List        :                            
'                             +pPI4G010ListMonthlyInvSvr
'*  7. Modified date(First) : 2000/10/16
'*  8. Modified date(Last)  : 2000/10/16
'*  9. Modifier (First)     : Nam Hoon Kim
'* 10. Modifier (Last)      : Nam Hoon Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
             
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")   

Call HideStatusWnd 

Err.Clear
On Error Resume Next

Dim pPI4G010            
Dim strMode             
Dim StrNextKey  
Dim LngMaxRow  
Dim LngRow
Dim PvArr
	Const C_SHEETMAXROWS_D = 100
	
Dim I1_ief_supplied_select_char

Dim I2_i_stock_req_temp
Const I401_I2_login_dt = 0
Const I401_I2_plant_cd = 1
Const I401_I2_mvmt_flag = 2
Const I401_I2_reference_number = 3

ReDim I2_i_stock_req_temp(I401_I2_reference_number)

Dim I3_b_item_cd
Dim I4_good_mvmt_workset_temp_timestamp

Dim I5_next_good_mvmt_workset
Const I401_I5_temp_timestamp = 0
Const I401_I5_qty = 1

ReDim I5_next_good_mvmt_workset(I401_I5_qty)
    
Dim E2_i_stock_req_temp
Const I401_E2_login_dt = 0
Const I401_E2_mvmt_flag = 1
Const I401_E2_reference_number = 2    

Dim E3_b_item_by_plant_ss_qty
Dim E4_i_onhand_stock_good_on_hand_qty

Dim E5_good_mvmt_workset
Const I401_E5_temp_timestamp = 0
Const I401_E5_qty = 1    

Dim EG1_export_group
Const I401_EG1_E1_good_mvmt_workset_qty = 0
Const I401_EG1_E2_i_stock_req_temp_mvmt_flag = 1
Const I401_EG1_E2_i_stock_req_temp_reference_number = 2
Const I401_EG1_E2_i_stock_req_temp_req_dt = 3
Const I401_EG1_E2_i_stock_req_temp_plan_dt = 4
Const I401_EG1_E2_i_stock_req_temp_remain_qty = 5
Const I401_EG1_E2_i_stock_req_temp_plan_qty = 6
Const I401_EG1_E2_i_stock_req_temp_tracking_no = 7
Const I401_EG1_E2_i_stock_req_temp_pur_mfg = 8
Const I401_EG1_E2_i_stock_req_temp_req_type = 9    

Dim E7_b_item
Const I401_E7_item_cd = 0
Const I401_E7_item_nm = 1
Const I401_E7_spec = 2
Const I401_E7_basic_unit = 3

	I2_i_stock_req_temp(I401_I2_plant_cd) = Request("txtPlantCd")
	I3_b_item_cd                          = Request("txtItemCd")
	I4_good_mvmt_workset_temp_timestamp   = UNIConvDate(Request("txtYyyyMmDd"))
	I1_ief_supplied_select_char = "N"
	 
	Set pPI4G010 = Server.CreateObject("PI4G010.cIListStkReq")
	    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End          
	End If    
	 
	Call pPI4G010.I_LIST_STOCK_REQ(gStrGlobalCollection, C_SHEETMAXROWS_D, _
				I1_ief_supplied_select_char, _
				I2_i_stock_req_temp, _
				I3_b_item_cd, _
				I4_good_mvmt_workset_temp_timestamp, _
				I5_next_good_mvmt_workset, _
				E2_i_stock_req_temp, _
				E3_b_item_by_plant_ss_qty, _
				E4_i_onhand_stock_good_on_hand_qty, _
				E5_good_mvmt_workset, _
				EG1_export_group, _
				E7_b_item)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPI4G010 = Nothing            
		Response.End             
	End If

	Set pPI4G010 = Nothing

	if isEmpty(EG1_export_group) then
		Call ServerMesgBox("조회된 결과가 없습니다.", vbInformation, I_MKSCRIPT)
		Response.End             
	End If
	 
	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	ReDim PvArr(ubound(EG1_export_group,1))

	For LngRow = 0 To ubound(EG1_export_group,1)
		
		strData = Chr(11) & UNIDateClientFormat(EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_req_dt))
		
		if EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_mvmt_flag) ="O" then
			strData = strData & Chr(11) & "출고"
		elseif EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_mvmt_flag) ="I" then
			strData = strData & Chr(11) & "입고"
		end if
		strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_tracking_no)) & _
							Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_plan_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
							Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I401_EG1_E2_i_stock_req_temp_remain_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
							Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I401_EG1_E1_good_mvmt_workset_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
							Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		PvArr(LngRow) = strData
	Next
		strData = Join(PvArr, "")
		
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If
	
	Response.Write "<Script Language=vbscript> " & vbCr   
	Response.Write " With Parent "               & vbCr
	Response.Write "   .ggoSpread.Source     = .frm1.vspdData "    & vbCr
	Response.Write "   .ggoSpread.SSShowData """ & strData  & """" & vbCr
	Response.Write "   .frm1.txtItemCd2.value  = """ & ConvSPChars(E7_b_item(I401_E7_item_cd))                                 & """" & vbCr
	Response.Write "   .frm1.txtItemNm2.value  = """ & ConvSPChars(E7_b_item(I401_E7_item_nm))                                 & """" & vbCr
	Response.Write "   .frm1.txtItemSpec.value  = """ & ConvSPChars(E7_b_item(I401_E7_spec))                                    & """" & vbCr
	Response.Write "   .frm1.txtBaseUnit.value  = """ & ConvSPChars(E7_b_item(I401_E7_basic_unit))                              & """" & vbCr
	Response.Write "   .frm1.txtOnhandQty.value = """ & UniConvNumberDBToCompany(E4_i_onhand_stock_good_on_hand_qty,ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & """" & vbCr
	Response.Write "   .frm1.txtSsQty.value  = """ & UniConvNumberDBToCompany(E3_b_item_by_plant_ss_qty,ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)          & """" & vbCr
	
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else										"				& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If										"				& vbCr
	
	Response.Write "End With       " & vbCr                    
	Response.Write "</Script>      " & vbCr   
	Response.End 
%>


