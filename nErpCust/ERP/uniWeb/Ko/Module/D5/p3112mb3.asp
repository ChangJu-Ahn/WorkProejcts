<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p3112mb3.asp
'*  4. Program Name         : Production Configuration (Query)
'*  5. Program Desc         :
'*  6. Component List       : PB0C101.cPLkUpPltConfig
'*  7. Modified date(First) : 2000/11/28
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Err.Clear
On Error Resume Next
Call HideStatusWnd

Dim pPB0C101                 '☆ : 조회용 Component Dll 사용 변수 
Dim I1_select_char, I2_plant_cd
Dim E1_p_plant_config, E2_b_plant, iStatusCodeOfPrevNext

' E1_p_plant_config
Const P007_E1_mps_time_fence_flg = 0
Const P007_E1_mrp_time_fence_flg = 1
Const P007_E1_inv_scrap_flg = 2
Const P007_E1_insp_inv_scrap_mthd = 3
Const P007_E1_mrp_list_mthd = 4
Const P007_E1_scrap_mthd = 5
Const P007_E1_rls_inv_chk_flg = 6
Const P007_E1_auto_rcpt_flg = 7
Const P007_E1_child_req_flg = 8
Const P007_E1_exss_rcpt_flg = 9
Const P007_E1_ord_close_mthd = 10
Const P007_E1_prev_opr_chk_flg = 11
Const P007_E1_bkup_cls_ord_flg = 12
Const P007_E1_prod_etc_mthd = 13
Const P007_E1_prod_etc_flg = 14
Const P007_E1_inv_etc_mthd = 15
Const P007_E1_inv_etc_flg = 16
Const P007_E1_opr_no_aux_bit = 17
Const P007_E1_max_load_in_sys = 18
Const P007_E1_max_load_size = 19
Const P007_E1_multi_site_flg = 20
Const P007_E1_read_data_flg = 21
Const P007_E1_write_thresh = 22
Const P007_E1_description = 23
Const P007_E1_time_zone = 24
Const P007_E1_prod_mthd = 25
Const P007_E1_prod_flg = 26
Const P007_E1_prod_string = 27
'2003-03-18 추가 START
Const P007_E1_MPS_METHOD1 = 28
Const P007_E1_DELIVERY_ORDER_FLG1 = 29
Const P007_E1_BOM_HISTORY_FLG1 = 30
'2003-03-18 추가 END
Const P007_E1_routing_lt_flg1 = 31
'2005-03-07 add
Const P007_E1_ENG_BOM_flg1 = 32
'2005-09-17 add
Const P007_E1_OPR_COST_flag1 = 33
'2006-04-10 add
Const P007_E1_BACKLOG_flag = 34
'2006-07-18 add
Const P007_E1_prod_child_mthd = 35
Const P007_E1_prod_rsc_mthd = 36


' E2_b_plant
Const P007_E2_plant_cd = 0
Const P007_E2_plant_nm = 1
    
If Request("txtPlantCd1") = "" Then          '⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

I1_select_char = Request("PrevNextFlg")
I2_plant_cd = Request("txtPlantCd1")
    
Set pPB0C101 = Server.CreateObject("PB0C101.cPLkUpPltConfig")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB0C101.P_LOOK_UP_PLANT_CONFIGURE_SVR(gStrGlobalCollection, I1_select_char, I2_plant_cd, _
                     E1_p_plant_config, E2_b_plant, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB0C101 = Nothing               '☜: Unload Component
	Response.End
End If

Set pPB0C101 = Nothing               '☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If
 
'-----------------------
'Result data display area
'----------------------- 
' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
	Response.Write ".txtPlantCd1.Value = """ & ConvSPChars(E2_b_plant(P007_E2_plant_cd)) & """" & vbCrLf '☆: Plant Code
	Response.Write ".txtPlantNm1.Value = """ & ConvSPChars(E2_b_plant(P007_E2_plant_nm)) & """" & vbCrLf '☆: Plant Name  
  
	If Trim(E1_p_plant_config(P007_E1_rls_inv_chk_flg)) = "1" Then
		Response.Write ".rdoRlsInvChkFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal1 = 1" & vbCrLf
	ElseIf Trim(E1_p_plant_config(P007_E1_rls_inv_chk_flg)) = "2" Then
		Response.Write ".rdoRlsInvChkFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal1 = 2" & vbCrLf
	Else
		Response.Write ".rdoRlsInvChkFlg3.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal1 = 3" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_auto_rcpt_flg) = "Y" Then
		Response.Write ".rdoAutoRcptFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal2 = 1" & vbCrLf
	Else
		Response.Write ".rdoAutoRcptFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal2 = 2" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_prev_opr_chk_flg) = "Y" Then
		Response.Write ".rdoPreOprChkFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal3 = 1" & vbCrLf
	Else
		Response.Write ".rdoPreOprChkFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal3 = 2" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_prod_etc_mthd) = "Y" Then
		Response.Write ".rdoProdEtcMthd1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal5 = 1" & vbCrLf
	Else
		Response.Write ".rdoProdEtcMthd2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal5 = 2" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_prod_flg) = "Y" Then
		Response.Write ".rdoProdFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal9 = 1" & vbCrLf
	Else
		Response.Write ".rdoProdFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal9 = 2" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_ord_close_mthd) = "Y" Then
		Response.Write ".rdoOrdCloseMthd1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal6 = 1" & vbCrLf
	Else
		Response.Write ".rdoOrdCloseMthd2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal6 = 2" & vbCrLf
	End If
	 
	If E1_p_plant_config(P007_E1_exss_rcpt_flg) = "Y" Then
		Response.Write ".rdoExssRcptFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal7 = 1" & vbCrLf
	Else
		Response.Write ".rdoExssRcptFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal7 = 2" & vbCrLf
	End If
	 
	Select Case E1_p_plant_config(P007_E1_prod_mthd)
		Case "1" 
			Response.Write ".rdoProdMthd1.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal8 = 1" & vbCrLf
		Case "2"
			Response.Write ".rdoProdMthd2.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal8 = 2" & vbCrLf
		Case "3"
			Response.Write ".rdoProdMthd3.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal8 = 3" & vbCrLf
		Case "4"
			Response.Write ".rdoProdMthd4.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal8 = 4" & vbCrLf
		Case "5"
			Response.Write ".rdoProdMthd5.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal8 = 5" & vbCrLf
	End Select
	 
	Response.Write ".txtProdEtcFlg.Value = """ & E1_p_plant_config(P007_E1_prod_etc_flg) & """" & vbCrLf 


'2003-03-18 추가 START	 
	Select Case E1_p_plant_config(P007_E1_MPS_METHOD1)
		Case "N" 
			Response.Write ".rdoMPSMETHOD1.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal10 = 1" & vbCrLf
		Case "S"
			Response.Write ".rdoMPSMETHOD2.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal10 = 2" & vbCrLf
		Case "C"
			Response.Write ".rdoMPSMETHOD3.Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal10 = 3" & vbCrLf
	End Select	 

	If E1_p_plant_config(P007_E1_DELIVERY_ORDER_FLG1) = "Y" Then
		Response.Write ".rdoDELIVERYORDERFLG1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal11 = 1" & vbCrLf
	Else
		Response.Write ".rdoDELIVERYORDERFLG2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal11 = 2" & vbCrLf
	End If	
	
	If E1_p_plant_config(P007_E1_BOM_HISTORY_FLG1) = "Y" Then
		Response.Write ".rdoBOMHISTORYFLG1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal12 = 1" & vbCrLf
	Else
		Response.Write ".rdoBOMHISTORYFLG2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal12 = 2" & vbCrLf
	End If	 	 
'2003-03-18 추가 END
	If E1_p_plant_config(P007_E1_routing_lt_flg1) = "Y" Then
		Response.Write ".rdoRoutingLTFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal13 = 1" & vbCrLf
	Else
		Response.Write ".rdoRoutingLTFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal13 = 2" & vbCrLf
	End If	 	 
	 
'2005-03-07 Add Start
	If E1_p_plant_config(P007_E1_ENG_BOM_flg1) = "Y" Then
		Response.Write ".rdoENGBOMFLG1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal14 = 1" & vbCrLf
	Else
		Response.Write ".rdoENGBOMFLG2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal14 = 2" & vbCrLf
	End If	 	 

'2005-03-07 Add End	 

'2005-09-17 Add Start
	If E1_p_plant_config(P007_E1_OPR_COST_flag1) = "Y" Then
		Response.Write ".rdoOprCostFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal15 = 1" & vbCrLf
	Else
		Response.Write ".rdoOprCostFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal15 = 2" & vbCrLf
	End If	 	 

'2005-09-17 Add End	 

'2006-04-17 Add Start
	If E1_p_plant_config(P007_E1_BACKLOG_flag) = "Y" Then
		Response.Write ".rdoBacklogFlg1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal16 = 1" & vbCrLf
	Else
		Response.Write ".rdoBacklogFlg2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal16 = 2" & vbCrLf
	End If	 	 

'2006-04-17 Add End	 


'2006-07-18 Add Start
Select Case E1_p_plant_config(P007_E1_prod_child_mthd)
	Case "1" 
		Response.Write ".rdoProdChildMthd1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal17 = 1" & vbCrLf
	Case "2"
		Response.Write ".rdoProdChildMthd2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal17 = 2" & vbCrLf
	Case "3"
		Response.Write ".rdoProdChildMthd3.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal17 = 3" & vbCrLf
	Case "4"
		Response.Write ".rdoProdChildMthd4.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal17 = 4" & vbCrLf
	Case "5"
		Response.Write ".rdoProdChildMthd5.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal17 = 5" & vbCrLf
End Select


Select Case E1_p_plant_config(P007_E1_prod_rsc_mthd)
	Case "1" 
		Response.Write ".rdoProdRscMthd1.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal18 = 1" & vbCrLf
	Case "2"
		Response.Write ".rdoProdRscMthd2.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal18 = 2" & vbCrLf
	Case "3"
		Response.Write ".rdoProdRscMthd3.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal18 = 3" & vbCrLf
	Case "4"
		Response.Write ".rdoProdRscMthd4.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal18 = 4" & vbCrLf
	Case "5"
		Response.Write ".rdoProdRscMthd5.Checked = True" & vbCrLf
		Response.Write "parent.lgRdoOldVal18 = 5" & vbCrLf
End Select

'2006-07-18 Add End
	 
	Response.Write "parent.DbQueryOk" & vbCrLf              '☜: 조화가 성공 
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Response.End                 '☜: Process End
%>

