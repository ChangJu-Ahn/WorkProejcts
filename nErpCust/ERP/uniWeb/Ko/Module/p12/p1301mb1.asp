<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%Call LoadBasisGlobalInf
  Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1301mb1.asp	
'*  4. Program Name         : Work Center
'*  5. Program Desc         :
'*  6. Component List       : PP1C201.cPLkWkCtr
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPP1C201																'☆ : 조회용 Component Dll 사용 변수 

Dim I1_select_char, I2_plant_cd, I3_wc_cd
Dim E1_b_plant, E2_p_mfg_calendar_type, E3_b_cost_center, E4_p_work_center, iStatusCodeOfPrevNext

Const P114_E1_plant_cd = 0
Const P114_E1_plant_nm = 1

Const P114_E2_cal_type = 0
Const P114_E2_cal_type_nm = 1

Const P114_E3_cost_cd = 0
Const P114_E3_cost_nm = 1

Const P114_E4_wc_cd = 0
Const P114_E4_wc_nm = 1
Const P114_E4_inside_flg = 2
Const P114_E4_wc_mgr = 3
Const P114_E4_valid_from_dt = 4
Const P114_E4_valid_to_dt = 5

If Request("txtPlantCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)         
	Response.End 
End If

If Request("txtConWcCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)                   
	Response.End 	
End If
	
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_select_char = Request("PrevNextFlg")
I2_plant_cd = Trim(Request("txtPlantCd"))
I3_wc_cd = Trim(Request("txtConWcCd"))
    
'-----------------------
'Com action area
'-----------------------
Set pPP1C201 = Server.CreateObject("PP1C201.cPLkUpWkCtr")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1C201.P_LOOK_UP_WORK_CENTER_SVR(gStrGlobalCollection, I1_select_char, I2_plant_cd, I3_wc_cd, _
                     E1_b_plant, E2_p_mfg_calendar_type, E3_b_cost_center, E4_p_work_center, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1C201 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
	Response.Write "	.txtPlantNm.value = """ & ConvSPChars(E1_b_plant(P114_E1_plant_nm)) & """" & vbCrLf		'☆: Plant Name
	Response.Write "	.txtConWcNm.value = """ & ConvSPChars(E4_p_work_center(P114_E4_wc_nm)) & """" & vbCrLf	'☆: Work Center Name
	Response.Write "	.txtConWcCd.Focus()" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

Set pPP1C201 = Nothing															'☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
Response.Write "	.txtPlantCd.value = """ & ConvSPChars(E1_b_plant(P114_E1_plant_cd)) & """" & vbCrLf		'☆: Plant Code
Response.Write "	.txtPlantNm.value = """ & ConvSPChars(E1_b_plant(P114_E1_plant_nm)) & """" & vbCrLf		'☆: Plant Name
Response.Write "	.txtConWcCd.value = """ & ConvSPChars(E4_p_work_center(P114_E4_wc_cd)) & """" & vbCrLf	'☆: Work Center Code
Response.Write "	.txtConWcNm.value = """ & ConvSPChars(E4_p_work_center(P114_E4_wc_nm)) & """" & vbCrLf	'☆: Work Center Name
			
Response.Write "	.txtDataWcCd.value = """ & ConvSPChars(E4_p_work_center(P114_E4_wc_cd)) & """" & vbCrLf	'☆: Work Center Code
Response.Write "	.txtDataWcNm.value = """ & ConvSPChars(E4_p_work_center(P114_E4_wc_nm)) & """" & vbCrLf	'☆: Work Center Name
Response.Write "	.cboInsideFlg.value = """ & E4_p_work_center(P114_E4_inside_flg) & """" & vbCrLf		'☆: Work Center Type
Response.Write "	.txtClnrType.value = """ & Trim(ConvSPChars(E2_p_mfg_calendar_type(P114_E2_cal_type))) & """" & vbCrLf	'☆: Calendar Type
Response.Write "	.txtClnrTypeNm.value = """ & Trim(ConvSPChars(E2_p_mfg_calendar_type(P114_E2_cal_type_nm))) & """" & vbCrLf	'☆: Calendar Type Nm
Response.Write "	.cboWCMgr.value = """ & E4_p_work_center(P114_E4_wc_mgr) & """" & vbCrLf			'☆: Work Center Manager
Response.Write "	.txtCostCd.value = """ & ConvSPChars(E3_b_cost_center(P114_E3_cost_cd)) & """" & vbCrLf	'☆: Cost Center
Response.Write "	.txtCostNm.value = """ & ConvSPChars(E3_b_cost_center(P114_E3_cost_nm)) & """" & vbCrLf	'☆: Cost Center Nm
		
Response.Write "	.txtValidFromDt.text = """ & UNIDateClientFormat(E4_p_work_center(P114_E4_valid_from_dt)) & """" & vbCrLf	'☆: Valid From Date
Response.Write "	.txtValidToDt.text = """ & UNIDateClientFormat(E4_p_work_center(P114_E4_valid_to_dt)) & """" & vbCrLf	'☆: Valid To Date
		
Response.Write "	parent.DbQueryOk" & vbCrLf																'☜: 조화가 성공 
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																	'☜: Process End
%>
