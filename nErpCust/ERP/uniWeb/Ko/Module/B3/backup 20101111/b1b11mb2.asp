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
'*  3. Program ID           : b1b11mb2.asp	
'*  4. Program Name         : Entry Item By Plant(Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PB3S107.cBMngItemByPlt
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3S107																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_b_item, I3_b_item_by_plant, I4_i_material_valuation, iCommandSent
Dim iIntFlgMode

' I2_b_item
Const P030_I2_item_cd = 0
Const P030_I2_basic_unit = 1
Const P030_I2_phantom_flg = 2

' I3_b_item_by_plant
Const P030_I3_procur_type = 0
Const P030_I3_item_acct = 1
Const P030_I3_single_rout_flg = 2
Const P030_I3_work_center = 3
Const P030_I3_valid_from_dt = 4
Const P030_I3_valid_to_dt = 5
Const P030_I3_order_unit_mfg = 6
Const P030_I3_order_lt_mfg = 7
Const P030_I3_order_lt_pur = 8
Const P030_I3_order_type = 9
Const P030_I3_order_rule = 10
Const P030_I3_round_perd = 11
Const P030_I3_prod_env = 12
Const P030_I3_mps_flg = 13
Const P030_I3_issue_mthd = 14
Const P030_I3_lot_flg = 15
Const P030_I3_cycle_cnt_perd = 16
Const P030_I3_major_sl_cd = 17
Const P030_I3_abc_flg = 18
Const P030_I3_recv_inspec_flg = 19
Const P030_I3_issued_sl_cd = 20
Const P030_I3_issued_unit = 21
Const P030_I3_order_unit_pur = 22
Const P030_I3_pur_org = 23
Const P030_I3_prod_inspec_flg = 24
Const P030_I3_final_inspec_flg = 25
Const P030_I3_ship_inspec_flg = 26
Const P030_I3_tracking_flg = 27
Const P030_I3_material_type = 28
Const P030_I3_order_from = 29
Const P030_I3_reorder_point = 30

' I4_i_material_valuation
Const P030_I4_prc_ctrl_indctr = 0
Const P030_I4_std_prc = 1
Const P030_I4_moving_avg_prc = 2

Redim I2_b_item(P030_I2_phantom_flg)
Redim I3_b_item_by_plant(P030_I3_reorder_point)
Redim I4_i_material_valuation(P030_I4_moving_avg_prc)

'-------------------------------------------------------------------------
' Validation Check
'-------------------------------------------------------------------------
If Request("txtPlantCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT) '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If

If Request("txtItemCd1") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT) '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
'-----------------------------------------
' 공장별 품목 일반 정보 
'-----------------------------------------

I1_plant_cd	= UCase(Trim(Request("txtPlantCd")))

I2_b_item(P030_I2_item_cd) = UCase(Trim(Request("txtItemCd1")))
I2_b_item(P030_I2_phantom_flg) = UCase(Trim(Request("txtPhantomFlg")))
I2_b_item(P030_I2_basic_unit) = UCase(Trim(Request("txtBasicUnit")))

I3_b_item_by_plant(P030_I3_item_acct) = UCase(Trim(Request("cboAccount")))
I3_b_item_by_plant(P030_I3_procur_type)	= UCase(Trim(Request("cboProcType")))
I3_b_item_by_plant(P030_I3_prod_env) = UCase(Trim(Request("cboProdEnv")))						 
I3_b_item_by_plant(P030_I3_material_type) = UCase(Trim(Request("cboMatType")))						 

If Request("rdoMPSItem") <> "" Then
	I3_b_item_by_plant(P030_I3_mps_flg) = UCase(Request("rdoMPSItem"))
Else
	I3_b_item_by_plant(P030_I3_mps_flg) = "N"
End If		
	
'------------------
'단공정여부필드 
'------------------
If Request("rdoCollectFlg") <> "" Then
	I3_b_item_by_plant(P030_I3_single_rout_flg) = UCase(Request("rdoCollectFlg"))
Else
	I3_b_item_by_plant(P030_I3_single_rout_flg) = "N"
End If
	
If Request("rdoTrackingItem") <> "" Then
	I3_b_item_by_plant(P030_I3_tracking_flg) = UCase(Request("rdoTrackingItem"))
Else
	I3_b_item_by_plant(P030_I3_tracking_flg) = "N"
End If		
			
I3_b_item_by_plant(P030_I3_work_center)	= UCase(Trim(Request("txtWorkcenter")))
	
If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 1, I_MKSCRIPT)
		Response.End	
	Else
		I3_b_item_by_plant(P030_I3_valid_from_dt) = UniConvDate(Request("txtValidFromDt"))	 
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 1, I_MKSCRIPT)
		Response.End	
	Else
		I3_b_item_by_plant(P030_I3_valid_to_dt) = UniConvDate(Request("txtValidToDt"))		 
	End If
End If
																										 
'-----------------------------------------
' 공장별MRP기준정보 
'-----------------------------------------
'==MRP 기준정보 
I3_b_item_by_plant(P030_I3_order_type) = UCase(Request("rdoMRPFlg"))
I3_b_item_by_plant(P030_I3_order_from) = UCase(Request("cboOrderFrom"))
I3_b_item_by_plant(P030_I3_reorder_point) = UniConvNum(Request("txtReorderPoint"), 0)

'==LOT SIZE정보 - 제조 
If Request("cboLotSizing") <> "" Then
	I3_b_item_by_plant(P030_I3_order_rule) = UCase(Request("cboLotSizing"))	
Else
	I3_b_item_by_plant(P030_I3_order_rule) = "L"
End If		
	
I3_b_item_by_plant(P030_I3_order_unit_mfg) = UCase(Request("txtMfgOrderUnit"))
I3_b_item_by_plant(P030_I3_order_lt_mfg) = UniCInt(Request("txtMfgOrderLT"), 0)		
I3_b_item_by_plant(P030_I3_round_perd) = UniCInt(Request("txtRoundPeriod"), 0)					 
	
'==LOT SIZE정보 - 구매 
I3_b_item_by_plant(P030_I3_order_unit_pur) = UCase(Request("txtPurOrderUnit"))					 
I3_b_item_by_plant(P030_I3_order_lt_pur) = UniCInt(Request("txtPurOrderLT"), 0)	
I3_b_item_by_plant(P030_I3_pur_org) = UCase(Request("txtPurOrg"))
	
'-----------------------------------------																 
' 재고 / 품질정보																						 
'-----------------------------------------
'==재고기준정보 
I3_b_item_by_plant(P030_I3_major_sl_cd) = UCase(Trim(Request("txtSLCd")))
	
If Trim(Request("cboIssueType")) = "" Then
	I3_b_item_by_plant(P030_I3_issue_mthd) = "A"
Else
	I3_b_item_by_plant(P030_I3_issue_mthd) = UCase(Trim(Request("cboIssueType")))
End If
	
I3_b_item_by_plant(P030_I3_issued_sl_cd) = UCase(Trim(Request("txtIssueSLCd")))
I3_b_item_by_plant(P030_I3_issued_unit) = UCase(Trim(Request("txtIssueUnit")))
I3_b_item_by_plant(P030_I3_lot_flg) = UCase(Request("rdoLotNoFlg"))					 
   
'Negative Stock여부 
I3_b_item_by_plant(P030_I3_cycle_cnt_perd) = UniConvNum(Request("txtCycleCntPerd"),0)				 
I3_b_item_by_plant(P030_I3_abc_flg) = UCase(Request("cboABCFlg"))						 

'==검사기준정보 
I3_b_item_by_plant(P030_I3_recv_inspec_flg) = UCase(Request("rdoPurInspType"))   '수입검사여부 
I3_b_item_by_plant(P030_I3_prod_inspec_flg) = UCase(Request("rdoMfgInspType"))   '공정검사여부 
I3_b_item_by_plant(P030_I3_final_inspec_flg) = UCase(Request("rdoFinalInspType")) '최종검사여부 
I3_b_item_by_plant(P030_I3_ship_inspec_flg) = UCase(Request("rdoIssueInspType")) '출하검사여부 

'==재고평가기준정보 

I4_i_material_valuation(P030_I4_prc_ctrl_indctr) = UCase(Request("cboPrcCtrlInd"))
I4_i_material_valuation(P030_I4_std_prc) = UniConvNum(Request("txtStdPrice"), 0)	
'Moving Average Price는 처음 Create Mode일 때만 입력 가능 
I4_i_material_valuation(P030_I4_moving_avg_prc) = UniConvNum(Request("txtMoveAvgPrice"), 0)
   
If iIntFlgMode = OPMD_CMODE Then																	 
	iCommandSent = "CREATE"						
ElseIf iIntFlgMode = OPMD_UMODE Then					
	iCommandSent = "UPDATE"						
End If

Set pPB3S107 = Server.CreateObject("PB3S107.cBMngItemByPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S107.B_MANAGE_ITEM_BY_PLANT(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_b_item, _
									I3_b_item_by_plant, I4_i_material_valuation)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S107 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB3S107 = Nothing															'☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>