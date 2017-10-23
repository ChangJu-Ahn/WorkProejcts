<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211MB1
'*  4. Program Name         : 품목별공급처등록 
'*  5. Program Desc         : 품목별공급처등록 
'*  6. Component List       : PM1S111.cMMaintSpplByItemS/PM1G219.cMLookupSpplByItemS/PB3S106.cBLkUpItemByPlt
'*  7. Modified date(First) : 2000/05/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin-hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
	Dim lgOpModeCRUD
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
			 Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case  "LookUpItemPlant"                                                                 '☜: Check	
             Call ChangeItemPlant()
    End Select


Sub SubBizSave()
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	Dim iPM1S111
	Dim import_m_supplier_item_by_plant
	Dim import_b_plant
	Dim import_b_item
	Dim import_b_biz_partner
	Dim import_b_pur_grp	
    Const M343_L1_pur_priority = 0
    Const M343_L1_sppl_dlvy_lt = 1
    Const M343_L1_def_flg = 2
    Const M343_L1_usage_flg = 3
    Const M343_L1_sppl_item_cd = 4
    Const M343_L1_sppl_item_nm = 5
    Const M343_L1_sppl_item_spec = 6
    Const M343_L1_maker_nm = 7
    Const M343_L1_valid_fr_dt = 8
    Const M343_L1_valid_to_dt = 9
    Const M343_L1_sppl_sales_prsn = 10
    Const M343_L1_sppl_tel_no = 11
    Const M343_L1_under_tol = 12
    Const M343_L1_over_tol = 13
    Const M343_L1_min_qty = 14
    Const M343_L1_max_qty = 15
    Const M343_L1_pur_unit = 16
    Const M343_L1_ext1_cd = 17
    Const M343_L1_ext1_qty = 18
    Const M343_L1_ext1_amt = 19
    Const M343_L1_ext2_cd = 20
    Const M343_L1_ext2_qty = 21
    Const M343_L1_ext2_amt = 22
    Const M343_L1_ext3_cd = 23
    Const M343_L1_ext3_qty = 24
    Const M343_L1_ext3_amt = 25
    
	Redim import_m_supplier_item_by_plant(M343_L1_ext3_amt)

    If Len(Trim(Request("txtfrdt"))) Then
		If UNIConvDate(Request("txtfrdt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtfrdt", 0, I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
	If Len(Trim(Request("txttodt"))) Then
		If UNIConvDate(Request("txttodt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txttodt", 0, I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
    import_m_supplier_item_by_plant(M343_L1_pur_priority) = UNIConvNum(Request("txtPriority"),0)
    import_m_supplier_item_by_plant(M343_L1_sppl_item_cd) = UCase(Trim(Request("txtSpplCd")))
    import_m_supplier_item_by_plant(M343_L1_sppl_item_nm) = Trim(Request("txtSpplNm"))
    import_m_supplier_item_by_plant(M343_L1_sppl_item_spec) = Trim(Request("txtSpplSpec"))
    import_m_supplier_item_by_plant(M343_L1_maker_nm) = Trim(Request("txtMakerNm"))
    if Trim(Request("txtFrDt")) <> "" then
		import_m_supplier_item_by_plant(M343_L1_valid_fr_dt) = UNIConvDate(Request("txtFrDt"))
    end if
    if Trim(Request("txtToDt")) <> "" then
		import_m_supplier_item_by_plant(M343_L1_valid_to_dt) = UNIConvDate(Request("txtToDt"))
    end if
    import_m_supplier_item_by_plant(M343_L1_usage_flg) = Trim(Request("rdoUseflg"))
    import_m_supplier_item_by_plant(M343_L1_sppl_sales_prsn) = Trim(Request("txtSpplPrsn"))
    import_m_supplier_item_by_plant(M343_L1_sppl_tel_no) = Trim(Request("txtTel"))
    import_m_supplier_item_by_plant(M343_L1_sppl_dlvy_lt) = UNIConvNum(Request("txtPurlt"),0)
    import_m_supplier_item_by_plant(M343_L1_under_tol) = UNIConvNUm(Request("txtUnder"),0)
    import_m_supplier_item_by_plant(M343_L1_over_tol) = UNIConvNum(Request("txtOver"),0)
    import_m_supplier_item_by_plant(M343_L1_min_qty) = UNIConvNum(Request("txtMinQty"),0)
    import_m_supplier_item_by_plant(M343_L1_def_flg) = UCase(Trim(Request("rdoDefFlg")))
    import_m_supplier_item_by_plant(M343_L1_max_qty) = UNIConvNum(Request("txtMaxQty"),0)
    import_m_supplier_item_by_plant(M343_L1_pur_unit) = UCase(Trim(Request("txtUnit")))
    import_m_supplier_item_by_plant(M343_L1_ext1_cd) = ""
    import_m_supplier_item_by_plant(M343_L1_ext1_qty) = 0
    import_m_supplier_item_by_plant(M343_L1_ext1_amt) = 0
    import_m_supplier_item_by_plant(M343_L1_ext2_cd) = ""
    import_m_supplier_item_by_plant(M343_L1_ext2_qty) = 0
    import_m_supplier_item_by_plant(M343_L1_ext2_amt) = 0	
    import_m_supplier_item_by_plant(M343_L1_ext3_cd) = ""
    import_m_supplier_item_by_plant(M343_L1_ext3_qty) = 0
    import_m_supplier_item_by_plant(M343_L1_ext3_amt) = 0
    
	lgOpModeCRUD = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	Set iPM1S111 = Server.CreateObject("PM1S111.cMMaintSpplByItemS")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM1S111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	import_b_plant		 = UCase(Trim(Request("txtPlantCd2")))
	import_b_item		 = UCase(Trim(Request("txtItemCd2")))
	import_b_biz_partner = UCase(Trim(Request("txtSupplierCd2")))
	import_b_pur_grp	 = UCase(Trim(Request("txtGroupCd")))

    If lgOpModeCRUD = OPMD_CMODE Then
		Call iPM1S111.M_MAINT_SPPL_BY_ITEM_SVR(gStrGlobalCollection, "CREATE", import_m_supplier_item_by_plant, _
							import_b_item, import_b_plant,import_b_pur_grp, import_b_biz_partner)
    ElseIf lgOpModeCRUD = OPMD_UMODE Then
		Call iPM1S111.M_MAINT_SPPL_BY_ITEM_SVR(gStrGlobalCollection, "UPDATE",	import_m_supplier_item_by_plant, _
				  import_b_item, import_b_plant, import_b_pur_grp, import_b_biz_partner )
    End If
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM1S111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM1S111 = Nothing
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "parent.frm1.txtPlantCd1.value		= """ & ConvSPChars(UCase(Trim(Request("txtPlantCd2"))))       & """" & vbCr
	Response.Write "parent.frm1.txtItemCd1.value		= """ & ConvSPChars(UCase(Trim(Request("txtItemCd2"))))        & """" & vbCr  
	Response.Write "parent.frm1.txtSupplierCd1.value	= """ & ConvSPChars(UCase(Trim(Request("txtSupplierCd2"))))    & """" & vbCr       
    Response.Write "Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "            
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	Err.Clear 

	Dim iPM1G219
	Dim import_b_biz_partner
	Dim import_b_item
	Dim import_b_plant
	Dim export_b_storage_location
    Const C_export_b_storage_location_sl_cd = 0
    Const C_export_b_storage_location_sl_nm = 1

	Dim export_m_supplier_item_by_plant
	Const C_e_b_plant_plant_cd = 0
	Const C_e_b_plant_plant_nm = 1
	Const C_e_b_item_item_cd = 2
	Const C_e_b_item_item_nm = 3
	Const C_e_b_biz_partner_bp_cd = 4
	Const C_e_b_biz_partner_bp_nm = 5
	Const C_e_b_pur_grp_pur_grp = 6
	Const C_e_b_pur_grp_pur_grp_nm = 7
	Const C_e_pur_priority = 8
    Const C_e_pur_unit = 9
    Const C_e_sppl_item_cd = 10
    Const C_e_sppl_item_nm = 11
    Const C_e_sppl_item_spec = 12
    Const C_e_maker_nm = 13
    Const C_e_valid_fr_dt = 14
    Const C_e_valid_to_dt = 15
    Const C_e_usage_flg = 16
    Const C_e_sppl_sales_prsn = 17
    Const C_e_sppl_tel_no = 18
    Const C_e_sppl_dlvy_lt = 19
    Const C_e_under_tol = 20
    Const C_e_over_tol = 21
    Const C_e_min_qty = 22
    Const C_e_def_flg = 23
    Const C_e_max_qty = 24
    Const C_e_ext1_cd = 25
    Const C_e_ext1_qty = 26
    Const C_e_ext1_amt = 27
    Const C_e_ext2_cd = 28
    Const C_e_ext2_qty = 29
    Const C_e_ext2_amt = 30
    Const C_e_ext3_cd = 31
    Const C_e_ext3_qty = 32
    Const C_e_ext3_amt = 33
    Const C_e_quota_rate = 35
    
    import_b_plant 	     = Trim(Request("txtPlantCd1"))
    import_b_item 		 = Trim(Request("txtItemCd1"))
    import_b_biz_partner = Trim(Request("txtSupplierCd1"))
    
    Set iPM1G219 = Server.CreateObject("PM1G219.cMLookupSpplByItemS")    
    
	If CheckSYSTEMError(Err,True) = true then 		
		Set iPM1G219 = Nothing
        Exit Sub
	End if

    Call iPM1G219.M_LOOKUP_SPPL_BY_ITEM_SVR(gStrGlobalCollection,import_b_biz_partner,import_b_item, import_b_plant, _
											export_b_storage_location, export_m_supplier_item_by_plant)
    
    If CheckSYSTEMError2(Err,True, "","","","","") = True Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr	
		Response.Write ".DbQueryOk1 " & vbCr	
		Response.Write "End With "				& vbCr   
		Response.Write "</Script> "	
		Set iPM1G219 = Nothing
        Exit Sub
	End if
	
	Set iPM1G219 = Nothing												'☜ 

	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write ".frm1.txtPlantCd1.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_plant_plant_cd))) & """" & vbCr
	Response.Write ".frm1.txtPlantNm1.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_plant_plant_nm))) & """" & vbCr
	Response.Write ".frm1.txtItemcd1.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_item_item_cd))) & """" & vbCr
	Response.Write ".frm1.txtItemNm1.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_item_item_nm))) & """" & vbCr
	Response.Write ".frm1.txtSupplierCd1.value		= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_biz_partner_bp_cd))) & """" & vbCr
	Response.Write ".frm1.txtSupplierNm1.value		= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_biz_partner_bp_nm))) & """" & vbCr
	Response.Write ".frm1.txtPlantCd2.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_plant_plant_cd))) & """" & vbCr
	Response.Write ".frm1.txtPlantNm2.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_plant_plant_nm))) & """" & vbCr
	Response.Write ".frm1.txtItemCd2.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_item_item_cd))) & """" & vbCr
	Response.Write ".frm1.txtItemNm2.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_item_item_nm))) & """" & vbCr
	Response.Write ".frm1.txtSupplierCd2.Value		= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_biz_partner_bp_cd))) & """" & vbCr
	Response.Write ".frm1.txtSupplierNm2.Value		= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_biz_partner_bp_nm))) & """" & vbCr
	Response.Write ".frm1.txtGroupCd.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_pur_grp_pur_grp))) & """" & vbCr
	Response.Write ".frm1.txtGroupNm.Value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_b_pur_grp_pur_grp_nm))) & """" & vbCr
	Response.Write ".frm1.txtPriority.text			= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_pur_priority), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtQuotaRate.text			= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_quota_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtPurLt.text				= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_sppl_dlvy_lt), ggQty.DecPoint, 0) & """" & vbCr
		
	If export_m_supplier_item_by_plant(C_e_usage_flg)="Y" then
		Response.Write ".frm1.rdoUseflg(0).checked = True " & vbcr
	Else
		Response.Write ".frm1.rdoUseflg(1).checked = True " & vbcr
	End If		
		
	If export_m_supplier_item_by_plant(C_e_def_flg)="Y" then
		Response.Write ".frm1.rdoDefFlg(0).checked = True " & vbcr
	Else
		Response.Write ".frm1.rdoDefFlg(1).checked = True " & vbcr
	End if
		
	Response.Write ".frm1.txtMinQty.text 			= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_min_qty), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtMaxQty.text 			= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_max_qty), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtUnit.Value				= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_pur_unit))) & """" & vbCr
	Response.Write ".frm1.txtStorageCd.Value		= """ & Trim(ConvSPChars(export_b_storage_location(C_export_b_storage_location_sl_cd))) & """" & vbCr
	Response.Write ".frm1.txtStorageNm.Value		= """ & Trim(ConvSPChars(export_b_storage_location(C_export_b_storage_location_sl_nm))) & """" & vbCr
	Response.Write ".frm1.txtOver.text				= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_over_tol), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtUnder.text				= """ & UNINumClientFormat(export_m_supplier_item_by_plant(C_e_under_tol), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".frm1.txtFrDt.Text				= """ & UNIDateClientFormat(export_m_supplier_item_by_plant(C_e_valid_fr_dt)) & """" & vbCr
	
	If CDate(UNIDateClientFormat(export_m_supplier_item_by_plant(C_e_valid_fr_dt))) >= CDate("2999-12-30") Then
	Response.Write ".frm1.txtToDt.Text				= """"" & vbCr
	Else
	Response.Write ".frm1.txtToDt.Text				= """ & UNIDateClientFormat(export_m_supplier_item_by_plant(C_e_valid_to_dt)) & """" & vbCr
	End If
	Response.Write ".frm1.txtSpplCd.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_sppl_item_cd))) & """" & vbCr
	Response.Write ".frm1.txtSpplNm.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_sppl_item_nm))) & """" & vbCr
	Response.Write ".frm1.txtSpplSpec.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_sppl_item_spec))) & """" & vbCr
	Response.Write ".frm1.txtMakerNm.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_maker_nm))) & """" & vbCr
	Response.Write ".frm1.txtSpplPrsn.value			= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_sppl_sales_prsn))) & """" & vbCr
	Response.Write ".frm1.txtTel.value				= """ & Trim(ConvSPChars(export_m_supplier_item_by_plant(C_e_sppl_tel_no))) & """" & vbCr
    Response.Write ".DbQueryOk "																		& vbCr   
    Response.Write "End With "																		& vbCr   
    Response.Write "</Script> "		
End Sub	

'============================================================================================================
' Name : SubBizDelete
' Desc : Date data 
'============================================================================================================
Sub SubBizDelete()
	Dim iPM1S111
	Dim import_b_plant
	Dim import_b_item
	Dim import_b_biz_partner
	Dim import_b_pur_grp	
	Dim import_m_supplier_item_by_plant
    
    On Error Resume Next
    Err.Clear

	Set iPM1S111 = Server.CreateObject("PM1S111.cMMaintSpplByItemS")
												
	If CheckSYSTEMError(Err,True) = True then 		
		Exit Sub
	End If
	
    import_b_plant			= UCase(Trim(Request("txtPlantCd1")))
    import_b_item			= UCase(Trim(Request("txtItemCd1")))
    import_b_biz_partner	= UCase(Trim(Request("txtSupplierCd1")))
    import_b_pur_grp		= UCase(Trim(Request("txtGroupCd")))
	
	Call iPM1S111.M_MAINT_SPPL_BY_ITEM_SVR(gStrGlobalCollection, "DELETE",	import_m_supplier_item_by_plant, _
				  import_b_item, import_b_plant, import_b_pur_grp, import_b_biz_partner )

	If CheckSYSTEMError(Err,True) = True Then
       Set iPM1S111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

	Set iPM1S111 = Nothing                                                                    '☜: Unload Comproxy
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write "Parent.DbDeleteOk "      & vbCr   
    Response.Write "</Script> "            
End Sub

'============================================================================================================
' Name : ChangeItemPlant
' Desc : 
'============================================================================================================
Sub ChangeItemPlant()
	Dim iB1b119
    On Error Resume Next                       '☜: Protect system from crashing
    Err.Clear								    '☜: Protect system	from crashing

    Dim I1_b_plant_plant_cd
    Dim I2_b_item_item_cd
    Dim E1_b_pur_org
    Const P003_E1_pur_org = 0
    Const P003_E1_pur_org_nm = 1
    Const P003_E1_valid_fr_dt = 2
    Const P003_E1_valid_to_dt = 3
    Const P003_E1_usage_flg = 4

    Dim E2_b_item_group
    Const P003_E2_item_group_cd = 0
    Const P003_E2_item_group_nm = 1
    Const P003_E2_leaf_flg = 2

    Dim E3_for_issued_b_storage_location
    Const P003_E3_sl_cd = 0
    Const P003_E3_sl_type = 1
    Const P003_E3_sl_nm = 2

    Dim E4_for_major_b_storage_location
    Const P003_E4_sl_cd = 0
    Const P003_E4_sl_type = 1
    Const P003_E4_sl_nm = 2

    Dim E5_i_material_valuation
    Const P003_E5_prc_ctrl_indctr = 0
    Const P003_E5_moving_avg_prc = 1
    Const P003_E5_std_prc = 2
    Const P003_E5_prev_std_prc = 3

    Dim E6_b_item_by_plant
    Const P003_E6_procur_type = 0
    Const P003_E6_order_unit_mfg = 1
    Const P003_E6_order_lt_mfg = 2
    Const P003_E6_order_lt_pur = 3
    Const P003_E6_order_type = 4
    Const P003_E6_order_rule = 5
    Const P003_E6_req_round_flg = 6
    Const P003_E6_fixed_mrp_qty = 7
    Const P003_E6_min_mrp_qty = 8
    Const P003_E6_max_mrp_qty = 9
    Const P003_E6_round_qty = 10
    Const P003_E6_round_perd = 11
    Const P003_E6_scrap_rate_mfg = 12
    Const P003_E6_ss_qty = 13
    Const P003_E6_prod_env = 14
    Const P003_E6_mps_flg = 15
    Const P003_E6_issue_mthd = 16
    Const P003_E6_mrp_mgr = 17
    Const P003_E6_inv_check_flg = 18
    Const P003_E6_lot_flg = 19
    Const P003_E6_cycle_cnt_perd = 20
    Const P003_E6_inv_mgr = 21
    Const P003_E6_major_sl_cd = 22
    Const P003_E6_abc_flg = 23
    Const P003_E6_mps_mgr = 24
    Const P003_E6_recv_inspec_flg = 25
    Const P003_E6_inspec_lt_mfg = 26
    Const P003_E6_inspec_mgr = 27
    Const P003_E6_valid_from_dt = 28
    Const P003_E6_valid_to_dt = 29
    Const P003_E6_item_acct = 30
    Const P003_E6_single_rout_flg = 31
    Const P003_E6_prod_mgr = 32
    Const P003_E6_issued_sl_cd = 33
    Const P003_E6_issued_unit = 34
    Const P003_E6_order_unit_pur = 35
    Const P003_E6_var_lt = 36
    Const P003_E6_scrap_rate_pur = 37
    Const P003_E6_pur_org = 38
    Const P003_E6_prod_inspec_flg = 39
    Const P003_E6_final_inspec_flg = 40
    Const P003_E6_ship_inspec_flg = 41
    Const P003_E6_inspec_lt_pur = 42
    Const P003_E6_option_flg = 43
    Const P003_E6_over_rcpt_flg = 44
    Const P003_E6_over_rcpt_rate = 45
    Const P003_E6_damper_flg = 46
    Const P003_E6_damper_max = 47
    Const P003_E6_damper_min = 48
    Const P003_E6_reorder_pnt = 49
    Const P003_E6_std_time = 50
    Const P003_E6_add_sel_rule = 51
    Const P003_E6_add_sel_value = 52
    Const P003_E6_add_seq_rule = 53
    Const P003_E6_add_seq_atrid = 54
    Const P003_E6_rem_sel_rule = 55
    Const P003_E6_rem_sel_value = 56
    Const P003_E6_rem_seq_rule = 57
    Const P003_E6_rem_seq_atrid = 58
    Const P003_E6_llc = 59
    Const P003_E6_tracking_flg = 60
    Const P003_E6_valid_flg = 61
    Const P003_E6_work_center = 62
    Const P003_E6_order_from = 63
    Const P003_E6_cal_type = 64
    Const P003_E6_line_no = 65
    Const P003_E6_atp_lt = 66
    Const P003_E6_etc_flg1 = 67
    Const P003_E6_etc_flg2 = 68

    Dim E7_b_item
    Const P003_E7_item_cd = 0
    Const P003_E7_item_nm = 1
    Const P003_E7_formal_nm = 2
    Const P003_E7_spec = 3
    Const P003_E7_item_acct = 4
    Const P003_E7_item_class = 5
    Const P003_E7_hs_cd = 6
    Const P003_E7_hs_unit = 7
    Const P003_E7_unit_weight = 8
    Const P003_E7_unit_of_weight = 9
    Const P003_E7_basic_unit = 10
    Const P003_E7_draw_no = 11
    Const P003_E7_item_image_flg = 12
    Const P003_E7_phantom_flg = 13
    Const P003_E7_blanket_pur_flg = 14
    Const P003_E7_base_item_cd = 15
    Const P003_E7_proportion_rate = 16
    Const P003_E7_valid_flg = 17
    Const P003_E7_valid_from_dt = 18
    Const P003_E7_valid_to_dt = 19

    Dim E8_b_plant
    Const P003_E8_plant_cd = 0
    Const P003_E8_plant_nm = 1
    
    Const strDefFrDate = "1899-12-31"
	Const strDefToDate = "2999-12-30"        
    
    Set	iB1b119 = CreateObject("PB3S106.cBLkUpItemByPlt")
   
    '-----------------------
    'Com action	result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iB1b119 = Nothing
		Exit Sub
	End if
     
	I1_b_plant_plant_cd = UCase(Trim(Request("txtPlantCd")))
    I2_b_item_item_cd	= UCase(Trim(Request("txtItemCd")))
           
    Call iB1b119.B_LOOK_UP_ITEM_BY_PLANT(gStrGlobalCollection, _
	   								I1_b_plant_plant_cd, _
	   								I2_b_item_item_cd, _
	   								E1_b_pur_org, _
	   								E2_b_item_group, _
	   								E3_for_issued_b_storage_location, _
	   								E4_for_major_b_storage_location, _
	   								E5_i_material_valuation, _
	   								E6_b_item_by_plant, _
	   								E7_b_item, _
	   								E8_b_plant)
 
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iB1b119 = Nothing									
		Exit Sub												
	
	End if
     
	'-----------------------
	'Result	data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.frm1.txtPlantCd2.value		= """ & Trim(ConvSPChars(E8_b_plant(P003_E8_plant_cd)))  & """" & vbCr
	Response.Write "Parent.frm1.txtPlantNm2.value		= """ & Trim(ConvSPChars(E8_b_plant(P003_E8_plant_nm)))  & """" & vbCr
	Response.Write "Parent.frm1.txtItemCd2.value		= """ & Trim(ConvSPChars(E7_b_item(P003_E7_item_cd)))  & """" & vbCr
	Response.Write "Parent.frm1.txtItemNm2.value		= """ & Trim(ConvSPChars(E7_b_item(P003_E7_item_nm)))  & """" & vbCr
	Response.Write "Parent.frm1.txtUnit.value			= """ & Trim(ConvSPChars(E6_b_item_by_plant(P003_E6_order_unit_pur)))  & """" & vbCr
	Response.Write "Parent.frm1.hdnOrg.value			= """ & Trim(ConvSPChars(E6_b_item_by_plant(P003_E6_pur_org)))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageCd.value		= """ & Trim(ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_cd)))  & """" & vbCr
	Response.Write "Parent.frm1.txtStorageNm.value		= """ & Trim(ConvSPChars(E4_for_major_b_storage_location(P003_E4_sl_nm)))  & """" & vbCr
	
	If CDate(UNIDateClientFormat(E6_b_item_by_plant(P003_E6_valid_from_dt))) > CDate(strDefFrDate) Then 
		Response.Write "Parent.frm1.txtFrDt.Text			= """ & UNIDateClientFormat(E6_b_item_by_plant(P003_E6_valid_from_dt)) & """" & vbCr
	Else 
		Response.Write "Parent.frm1.txtFrDt.Text			= """"" & vbCr
	End If

	If CDate(UNIDateClientFormat(E6_b_item_by_plant(P003_E6_valid_to_dt))) > CDate(strDefToDate) Then 
		Response.Write "Parent.frm1.txtToDt.Text			= """"" & vbCr
	Else 
		Response.Write "Parent.frm1.txtToDt.Text			= """ & UNIDateClientFormat(E6_b_item_by_plant(P003_E6_valid_to_dt)) & """" & vbCr
	End If
	Response.Write "</Script>" & vbCr
End Sub	

%>


