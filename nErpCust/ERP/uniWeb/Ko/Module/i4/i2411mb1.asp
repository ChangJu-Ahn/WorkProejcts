<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Rop
'*  3. Program ID           : I2411mb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : ROP품목을 조회한다 
'*  6. Comproxy List        :                            
'                             +i24118ListMonthlyInvSvr
'*  7. Modified date(First) : 2000/05/03
'*  8. Modified date(Last)  : 2000/05/03
'*  9. Modifier (First)     : Soon Ho Kweon
'* 10. Modifier (Last)      : Soon Ho Kweon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/03 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!--<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>-->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")   

'Hide Processin Bar
Call HideStatusWnd 

On Error Resume Next
Err.Clear

Dim i24118            
Dim strMode            

Dim StrNextKey  
Dim LngMaxRow  
Dim LngRow
Dim PvArr

	Const C_SHEETMAXROWS_D = 500

    '-----------------------
    'IMPORTS View
    '-----------------------
    Dim I1_good_mvmt_workset_temp_timestamp
    Dim I2_b_item_cd
    Dim I3_b_plant_cd
    Dim I4_i_onhand_stk_tracking_no
        
 '-----------------------
 'EXPORTS View
 '-----------------------
    Dim EG1_export_group 
  Const I405_EG1_E1_b_item_item_cd = 0
  Const I405_EG1_E1_b_item_item_nm = 1
  Const I405_EG1_E1_b_item_spec = 2
  Const I405_EG1_E1_b_item_basic_unit = 3
  Const I405_EG1_E2_i_onhand_stock_good_on_hand_qty = 4
  Const I405_EG1_E3_b_item_by_plant_ss_qty = 5
  Const I405_EG1_E3_b_item_by_plant_reorder_pnt = 6
  Const I405_EG1_E3_b_item_by_plant_order_lt_pur = 7
  Const I405_EG1_E3_b_item_by_plant_fixed_mrp_qty = 8
  Const I405_EG1_E4_item_tot_issued_qty_good_mvmt_workset_qty = 9
  Const I405_EG1_E5_item_tot_schd_rcp_qty_good_mvmt_workset_qty = 10
  Const I405_EG1_E6_good_mvmt_workset_qty = 11
  Const I405_EG1_E2_i_onhand_stock_tracking_no = 12
    Dim E3_Next_b_item_cd


    StrNextKey = Request("lgStrPrevKey")
 
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I3_b_plant_cd  = Request("txtPlantCd")        
    I2_b_item_cd  = Request("txtItemCd")
    If Request("txtTrackingNo") = "" Then
		I4_i_onhand_stk_tracking_no = "*"
	Else
		I4_i_onhand_stk_tracking_no = Request("txtTrackingNo")
	End If
	
    if StrNextKey <> "" then I2_b_item_cd = StrNextKey
    
 
 Set i24118 = Server.CreateObject("PI4G030.cIListRopItemSvr")
    
 If CheckSYSTEMError(Err, True) = True Then
  Response.End            '☜: 비지니스 로직 처리를 종료함 
 End If    
 
 Call i24118.I_LIST_ROP_ITEM_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
         I1_good_mvmt_workset_temp_timestamp, _
         I2_b_item_cd, _
         I3_b_plant_cd, _
         I4_i_onhand_stk_tracking_no, _
         EG1_export_group, _
         E3_Next_b_item_cd)
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
		Set i24118 = Nothing           
		Response.End              
	End If

 Set i24118 = Nothing

 if isEmpty(EG1_export_group) then
	Response.End             
 End If

 
 strData = ""
 LngMaxRow = CLng(Request("txtMaxRows")) + 1
 ReDim PvArr(ubound(EG1_export_group,1))

	For LngRow = 0 To ubound(EG1_export_group,1)
		strData =	Chr(11) & "0" & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E1_b_item_item_cd)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E1_b_item_item_nm)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E1_b_item_spec)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E1_b_item_basic_unit)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E2_i_onhand_stock_tracking_no)) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E6_good_mvmt_workset_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E3_b_item_by_plant_reorder_pnt),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & "" & _
					Chr(11) & "" & _
					Chr(11) & "" & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E2_i_onhand_stock_good_on_hand_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow,I405_EG1_E5_item_tot_schd_rcp_qty_good_mvmt_workset_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E4_item_tot_issued_qty_good_mvmt_workset_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E3_b_item_by_plant_ss_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I405_EG1_E3_b_item_by_plant_fixed_mrp_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I405_EG1_E3_b_item_by_plant_order_lt_pur)) & _
					Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		PvArr(LngRow) = strData
	Next
		strData = Join(PvArr, "")
    
 If CheckSYSTEMError(Err, True) = True Then
  Response.End
 End If

 If E3_Next_b_item_cd = EG1_export_group(ubound(EG1_export_group,1), I405_EG1_E1_b_item_item_cd)  then
  StrNextKey = ""
 Else
  StrNextKey = E3_Next_b_item_cd
 End If   



    Response.Write "<Script Language=vbscript> " & vbCr   
    
    Response.Write " With Parent "               & vbCr
    
    Response.Write "   .ggoSpread.Source     = .frm1.vspdData "    & vbCr
    Response.Write "   .ggoSpread.SSShowData """ & strData  & """" & vbCr
      'Lock 처리 
    Response.Write "   .lgStrPrevKey  = """ & ConvSPChars(StrNextKey)                         & """" & vbCr
    Response.Write "   If.frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) and .lgStrPrevKey <> """"  Then " & vbCr
    Response.Write "		.DbQuery "                                                              & vbCr
    Response.Write "   else "                                                                       & vbCr
	Response.Write "		.frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd"))			& """" & vbCr
	Response.Write "		.frm1.hItemCd.value  = """ & ConvSPChars(Request("txtItemCd"))			& """" & vbCr
	Response.Write "		.frm1.hTrackingNo.value  = """ & ConvSPChars(Request("txtTrackingNo"))  & """" & vbCr
    Response.Write "    .DbQueryOk "																& vbCr  
    Response.Write "   end if  "                                                                    & vbCr  
    
    Response.Write " End With  " & vbCr                    
    Response.Write " </Script> " & vbCr   
    
    Response.End 
%>


