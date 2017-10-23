<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%Call LoadBasisGlobalInf
  Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  b1b11mb1.asp
'*  4. Program Name         :  Item by Plant 조회 
'*  5. Program Desc         :
'*  6. Component List       : PB3S106.cBLkUpItemByPlt
'*  7. Modified date(First) : 2000/04/21
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Dim pPB3S106																'☆ : 조회용 Component Dll 사용 변수 
Dim I1_select_char, I2_plant_cd, I3_item_cd 
Dim E5_i_material_valuation, E6_b_plant, E7_b_item, E8_b_item_by_plant, iStatusCodeOfPrevNext

' E5_i_material_valuation
Const P027_E5_prc_ctrl_indctr = 0
Const P027_E5_moving_avg_prc = 1
Const P027_E5_std_prc = 2
Const P027_E5_prev_std_prc = 3

' E6_b_plant
Const P027_E6_plant_cd = 0
Const P027_E6_plant_nm = 1

' E7_b_item
Const P027_E7_item_cd = 0
Const P027_E7_item_nm = 1
Const P027_E7_basic_unit = 10
Const P027_E7_phantom_flg = 13

' E8_b_item_by_plant
Const P027_E8_procur_type = 0
Const P027_E8_order_unit_mfg = 1
Const P027_E8_order_lt_mfg = 2
Const P027_E8_order_lt_pur = 3
Const P027_E8_order_type = 4
Const P027_E8_order_rule = 5
Const P027_E8_round_perd = 11
Const P027_E8_prod_env = 14
Const P027_E8_mps_flg = 15
Const P027_E8_issue_mthd = 16
Const P027_E8_lot_flg = 19
Const P027_E8_cycle_cnt_perd = 20
Const P027_E8_major_sl_cd = 22
Const P027_E8_abc_flg = 23
Const P027_E8_recv_inspec_flg = 25
Const P027_E8_valid_from_dt = 28
Const P027_E8_valid_to_dt = 29
Const P027_E8_item_acct = 30
Const P027_E8_single_rout_flg = 31
Const P027_E8_issued_sl_cd = 33
Const P027_E8_issued_unit = 34
Const P027_E8_order_unit_pur = 35
Const P027_E8_pur_org = 38
Const P027_E8_prod_inspec_flg = 39
Const P027_E8_final_inspec_flg = 40
Const P027_E8_ship_inspec_flg = 41
Const P027_E8_option_flg = 43
Const P027_E8_reorder_pnt = 49
Const P027_E8_tracking_flg = 60
Const P027_E8_work_center = 62
Const P027_E8_order_from = 63
Const P027_E8_material_type = 69

If Request("txtPlantCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If
    
If Request("txtItemCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT) 
	Response.End 
End If

'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_select_char = Request("PrevNextFlg")
I2_plant_cd = Request("txtPlantCd")
I3_item_cd = Request("txtItemCd")

Set pPB3S106 = Server.CreateObject("PB3S106.cBLkUpItemByPlt")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S106.B_LOOK_UP_ITEM_BY_PLANT_SVR(gStrGlobalCollection, I1_select_char, I2_plant_cd, I3_item_cd, , , , , _
                     E5_i_material_valuation, E6_b_plant, E7_b_item, E8_b_item_by_plant, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S106 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
	Response.Write ".txtPlantNm.value = """ & ConvSPChars(E6_b_plant(P027_E6_plant_nm)) & """" & vbCrLf
	Response.Write ".txtItemNm.value = """ & ConvSPChars(E7_b_item(P027_E7_item_nm)) & """" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Response.End
End If

Set pPB3S106 = Nothing															'☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If

'-----------------------
'Result data display area
'----------------------- 
Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf

'------------------------------------------					
' 공장별 품목 일반정보 
'------------------------------------------
Response.Write ".txtPlantCd.value = """ & Trim(ConvSPChars(E6_b_plant(P027_E6_plant_cd))) & """" & vbCrLf								
Response.Write ".txtPlantNm.value = """ & ConvSPChars(E6_b_plant(P027_E6_plant_nm)) & """" & vbCrLf
Response.Write ".txtItemCd.value = """ & Trim(ConvSPChars(E7_b_item(P027_E7_item_cd))) & """" & vbCrLf
Response.Write ".txtItemNm.value = """ & ConvSPChars(E7_b_item(P027_E7_item_nm)) & """" & vbCrLf

Response.Write ".txtItemCd1.value = """ & ConvSPChars(E7_b_item(P027_E7_item_cd)) & """" & vbCrLf
Response.Write ".txtItemNm1.value = """ & ConvSPChars(E7_b_item(P027_E7_item_nm)) & """" & vbCrLf
Response.Write ".cboAccount.value = """ & E8_b_item_by_plant(P027_E8_item_acct) & """" & vbCrLf
Response.Write ".cboProcType.value = """ & E8_b_item_by_plant(P027_E8_procur_type) & """" & vbCrLf
Response.Write ".cboMatType.value = """ & E8_b_item_by_plant(P027_E8_material_type) & """" & vbCrLf
Response.Write ".cboProdEnv.value = """ & E8_b_item_by_plant(P027_E8_prod_env) & """" & vbCrLf
		
If E8_b_item_by_plant(P027_E8_mps_flg) = "Y" Then
	Response.Write ".rdoMPSItem1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoMpsOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoMPSItem2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoMpsOldVal = 2" & vbCrLf
End If
		
If E8_b_item_by_plant(P027_E8_tracking_flg) = "Y" Then
	Response.Write ".rdoTrackingItem1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoTrkOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoTrackingItem2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoTrkOldVal = 2" & vbCrLf
End If

'--------------------------------------------
'Collective Flag를 단공정 여부 필드로 사용 
'--------------------------------------------
If E8_b_item_by_plant(P027_E8_single_rout_flg)  = "Y" Then
	Response.Write ".rdoCollectFlg1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoColOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoCollectFlg2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoColOldVal = 2" & vbCrLf
End If

Response.Write ".txtWorkCenter.value = """ & ConvSPChars(E8_b_item_by_plant(P027_E8_work_center)) & """" & vbCrLf
Response.Write ".txtValidFromDt.text = """ & UniDateClientFormat(E8_b_item_by_plant(P027_E8_valid_from_dt)) & """" & vbCrLf
Response.Write ".txtValidToDt.text = """ & UniDateClientFormat(E8_b_item_by_plant(P027_E8_valid_to_dt)) & """" & vbCrLf
		
'------------------------------------------					
' 공장별MRP기준정보 
'------------------------------------------	
'==== MRP 기준정보 - 제조				
If E8_b_item_by_plant(P027_E8_order_type) = "Y" Then
	Response.Write ".rdoMRPFlg1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoMrpOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoMRPFlg2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoMrpOldVal = 2" & vbCrLf
End If

Response.Write ".cboOrderFrom.value	= """ & E8_b_item_by_plant(P027_E8_order_from) & """" & vbCrLf
		
'==== LOT SIZE정보 - 제조 
Response.Write ".cboLotSizing.value	= """ & E8_b_item_by_plant(P027_E8_order_rule) & """" & vbCrLf
		
Response.Write ".txtMfgOrderUnit.value = """ & ConvSPChars(E8_b_item_by_plant(P027_E8_order_unit_mfg)) & """" & vbCrLf
Response.Write ".txtMfgOrderLT.Text =	""" & UniConvNumDBToCompanyWithOutChange(E8_b_item_by_plant(P027_E8_order_lt_mfg), 0) & """" & vbCrLf
Response.Write ".txtRoundPeriod.Text =	""" & UniConvNumDBToCompanyWithOutChange(E8_b_item_by_plant(P027_E8_round_perd), 0) & """" & vbCrLf
Response.Write ".txtReorderPoint.Text =	""" & UniConvNumberDBToCompany(E8_b_item_by_plant(P027_E8_reorder_pnt), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf	

'==== Lot Size 정보 - 구매 
Response.Write ".txtPurOrderUnit.value = """ & ConvSPChars(E8_b_item_by_plant(P027_E8_order_unit_pur)) & """" & vbCrLf
Response.Write ".txtPurOrderLT.Text =	""" & UniConvNumDBToCompanyWithOutChange(E8_b_item_by_plant(P027_E8_order_lt_pur), 0) & """" & vbCrLf
Response.Write ".txtPurOrg.value = """ & ConvSPChars(E8_b_item_by_plant(P027_E8_pur_org)) & """" & vbCrLf
		
'------------------------------------------					
' 공장별재고/품질정보 - TAB3
'------------------------------------------					
'==== 재고기준정보 
Response.Write ".txtSLCd.value = """ & ConvSPChars(E8_b_item_by_plant(P027_E8_major_sl_cd)) & """" & vbCrLf
Response.Write ".cboIssueType.value	= """ & ConvSPChars(E8_b_item_by_plant(P027_E8_issue_mthd)) & """" & vbCrLf
Response.Write ".txtIssueSLCd.value	= """ & ConvSPChars(E8_b_item_by_plant(P027_E8_issued_sl_cd)) & """" & vbCrLf
Response.Write ".txtIssueUnit.value	= """ & ConvSPChars(E8_b_item_by_plant(P027_E8_issued_unit)) & """" & vbCrLf

If E8_b_item_by_plant(P027_E8_lot_flg) = "Y" Then
	Response.Write ".rdoLotNoFlg1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoLotOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoLotNoFlg2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoLotOldVal = 2" & vbCrLf
End If
		
Response.Write ".txtCycleCntPerd.Text =	""" & UniConvNumDBToCompanyWithOutChange(E8_b_item_by_plant(P027_E8_cycle_cnt_perd), 0) & """" & vbCrLf
Response.Write ".cboABCFlg.value = """ & E8_b_item_by_plant(P027_E8_abc_flg) & """" & vbCrLf
		
'==== 검사기준정보 
If E8_b_item_by_plant(P027_E8_recv_inspec_flg) = "Y" Then
	Response.Write ".rdoPurInspType1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoRecOldVal	= 1" & vbCrLf
Else
	Response.Write ".rdoPurInspType2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoRecOldVal	= 2" & vbCrLf
End If

If E8_b_item_by_plant(P027_E8_prod_inspec_flg) = "Y" Then
	Response.Write ".rdoMfgInspType1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoPrdOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoMfgInspType2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoPrdOldVal = 2" & vbCrLf
End If

If E8_b_item_by_plant(P027_E8_final_inspec_flg) = "Y" Then
	Response.Write ".rdoFinalInspType1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoFinOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoFinalInspType2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoFinOldVal = 2" & vbCrLf
End If

If E8_b_item_by_plant(P027_E8_ship_inspec_flg) = "Y" Then
	Response.Write ".rdoIssueInspType1.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoIssOldVal = 1" & vbCrLf
Else
	Response.Write ".rdoIssueInspType2.Checked = True" & vbCrLf
	Response.Write "parent.lgRdoIssOldVal = 2" & vbCrLf
End If		

'==== 재고평가기준정보 
Response.Write ".cboPrcCtrlInd.value = """ & E5_i_material_valuation(P027_E5_prc_ctrl_indctr) & """" & vbCrLf
Response.Write ".txtStdPrice.Text = """ & UniConvNumberDBToCompany(E5_i_material_valuation(P027_E5_std_prc), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) & """" & vbCrLf
Response.Write ".txtPrevStdPrice.Text = """ & UniConvNumberDBToCompany(E5_i_material_valuation(P027_E5_prev_std_prc), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) & """" & vbCrLf
Response.Write ".txtMoveAvgPrice.Text = """ & UniConvNumberDBToCompany(E5_i_material_valuation(P027_E5_moving_avg_prc), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0) & """" & vbCrLf
		
Response.Write ".txtPhantomFlg.value = """ & E7_b_item(P027_E7_phantom_flg) & """" & vbCrLf
Response.Write ".txtBasicUnit.value	= """ & ConvSPChars(E7_b_item(P027_E7_basic_unit)) & """" & vbCrLf
		
Response.Write "parent.DbQueryOk" & vbCrLf	'☜: 조회가 성공 
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Response.End								'☜: Process End
%>