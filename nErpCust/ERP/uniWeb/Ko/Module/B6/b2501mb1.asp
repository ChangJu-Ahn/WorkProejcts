<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%Call LoadBasisGlobalInf
  Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b2501mb1.asp
'*  4. Program Name         : Entry Plant (Query)
'*  5. Program Desc         :
'*  6. Component List       : PB6S101.cBLkUpPlt
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB6S101																'☆ : 조회용 Component Dll 사용 변수 

Dim I1_select_char, I2_plant_cd
Dim E1_b_biz_area, E3_b_plant, E4_p_mfg_calendar_type, iStatusCodeOfPrevNext

Const P058_E1_biz_area_cd = 0
Const P058_E1_biz_area_nm = 1
Const P058_E1_biz_area_eng_nm = 2

Const P058_E3_plant_cd = 0
Const P058_E3_plant_nm = 1
Const P058_E3_cur_cd = 2
Const P058_E3_plan_hrzn = 3
Const P058_E3_dtf_for_mps = 7
Const P058_E3_ptf_for_mps = 8
Const P058_E3_ptf_for_mrp = 9
Const P058_E3_inv_cls_dt = 10
Const P058_E3_inv_open_dt = 11
Const P058_E3_valid_from_dt = 12
Const P058_E3_valid_to_dt = 13
Const P058_E3_country_cd = 14
Const P058_E3_country_nm = 15
Const P058_E3_s_o_flag = 16

Const P058_E4_cal_type = 0
Const P058_E4_cal_type_nm = 1

If Request("txtPlantCd1") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbInformation, "", "", I_MKSCRIPT)	
	Response.End 
End If
    
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_select_char = Request("PrevNextFlg")
I2_plant_cd = UCase(Trim(Request("txtPlantCd1")))
    
'-----------------------
'Com action area
'-----------------------
Set pPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")

If CheckSYSTEMError(Err, True) = True Then
	Response.End
End If

Call pPB6S101.B_LOOK_UP_PLANT_SVR(gStrGlobalCollection, I1_select_char, I2_plant_cd, E1_b_biz_area, , _
                     E3_b_plant, E4_p_mfg_calendar_type, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB6S101 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm1.value = """"" & vbCrLf				  '☆: Plant Name		
	Response.Write "parent.frm1.txtPlantCd1.Focus()" & vbCrLf				  '☆: Plant Name		
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

Set pPB6S101 = Nothing															'☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If

'-----------------------
'Result data display area
'----------------------- 
' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.

Response.Write "<Script Language=vbscript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
	Response.Write ".txtPlantCd1.value = """ & ConvSPChars(E3_b_plant(P058_E3_plant_cd)) & """" & vbCrLf				  '☆: Plant Code
	Response.Write ".txtPlantNm1.value = """ & ConvSPChars(E3_b_plant(P058_E3_plant_nm)) & """" & vbCrLf				  '☆: Plant Name		
	Response.Write ".txtPlantCd2.value = """ & ConvSPChars(E3_b_plant(P058_E3_plant_cd)) & """" & vbCrLf				  '☆: Plant Code
	Response.Write ".txtPlantNm2.value = """ & ConvSPChars(E3_b_plant(P058_E3_plant_nm)) & """" & vbCrLf				  '☆: Plant Name
	Response.Write ".txtBizAreaCd.value = """ & ConvSPChars(E1_b_biz_area(P058_E1_biz_area_cd)) & """" & vbCrLf			  '☆: BizArea Code
	Response.Write ".txtBizAreaNm.value = """ & ConvSPChars(E1_b_biz_area(P058_E1_biz_area_nm)) & """" & vbCrLf			  '☆: BizArea Name
	Response.Write ".txtClnrType.value = """ & ConvSPChars(E4_p_mfg_calendar_type(P058_E4_cal_type)) & """" & vbCrLf	  '☆: Calendar Type	
	Response.Write ".txtClnrTypeNm.value = """ & ConvSPChars(E4_p_mfg_calendar_type(P058_E4_cal_type_nm)) & """" & vbCrLf '☆: Calendar Type Name
	Response.Write ".txtCurCd.value = """ & ConvSPChars(E3_b_plant(P058_E3_cur_cd)) & """" & vbCrLf						  '☆: Currency Code
	Response.Write ".txtPlngHrzn.Text = """ & UniConvNumDBToCompanyWithOutChange(E3_b_plant(P058_E3_plan_hrzn), 0) & """" & vbCrLf 
	Response.Write ".txtDtfForMps.Text = """ & E3_b_plant(P058_E3_dtf_for_mps) & """" & vbCrLf							  '☆: DTF For MPS
	Response.Write ".txtPtfForMps.Text = """ & E3_b_plant(P058_E3_ptf_for_mps) & """" & vbCrLf							  '☆: PTF For MPS
	Response.Write ".txtPtfForMrp.Text = """ & E3_b_plant(P058_E3_ptf_for_mrp) & """" & vbCrLf							  '☆: DTF For MRP
	Response.Write ".txtInvClsDt.text = """ & UNIMonthClientFormat(E3_b_plant(P058_E3_inv_cls_dt)) & """" & vbCrLf		  '☆: Inv Closing Date(년월까지)
	Response.Write ".txtInvOpenDt.text = """ & UNIMonthClientFormat(E3_b_plant(P058_E3_inv_open_dt)) & """" & vbCrLf	  '☆: Inv Open Date(년월까지)
	Response.Write ".hInvClsDt.Value = """ & UNIDateClientFormat(E3_b_plant(P058_E3_inv_cls_dt)) & """" & vbCrLf		  '☆: Inv Closing Date(년월까지)
	Response.Write ".hInvOpenDt.Value = """ & UNIDateClientFormat(E3_b_plant(P058_E3_inv_open_dt)) & """" & vbCrLf		  '☆: Inv Open Date(년월까지)
	Response.Write ".txtValidFromDt.text = """ & UNIDateClientFormat(E3_b_plant(P058_E3_valid_from_dt)) & """" & vbCrLf	  '☆: Valid From Date
	Response.Write ".txtValidToDt.text = """ & UNIDateClientFormat(E3_b_plant(P058_E3_valid_to_dt)) & """" & vbCrLf		  '☆: Valid To Date
	Response.Write ".txtCountryCd.value = """ & ConvSPChars(E3_b_plant(P058_E3_country_cd)) & """" & vbCrLf	  '☆: Calendar Type	
	Response.Write ".txtCountryNm.value = """ & ConvSPChars(E3_b_plant(P058_E3_country_nm)) & """" & vbCrLf '☆: Calendar Type Name
	Response.Write ".cboSOFlag.value = """ & ConvSPChars(E3_b_plant(P058_E3_s_o_flag)) & """" & vbCrLf 
		
	Response.Write "parent.DbQueryOk" & vbCrLf																'☜: 조화가 성공 
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																	'☜: Process End
%>
