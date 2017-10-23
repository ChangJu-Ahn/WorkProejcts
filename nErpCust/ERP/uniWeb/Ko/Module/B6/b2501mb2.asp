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
'*  3. Program ID           : b2501mb2.asp	
'*  4. Program Name         : Entry Plant (Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PB6G102.cBMngPlt
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB6G102																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_b_plant, I2_biz_area_cd, I3_cal_type, iCommandSent
Dim iIntFlgMode

' I1_b_plant
Const P062_I1_plant_cd = 0
Const P062_I1_plant_nm = 1
Const P062_I1_cur_cd = 2
Const P062_I1_plan_hrzn = 3
Const P062_I1_dtf_for_mps = 4
Const P062_I1_ptf_for_mps = 5
Const P062_I1_ptf_for_mrp = 6
Const P062_I1_inv_cls_dt = 7
Const P062_I1_inv_open_dt = 8
Const P062_I1_valid_from_dt = 9
Const P062_I1_valid_to_dt = 10
Const P062_I1_country_cd = 11
Const P062_I1_s_o_flag = 12

Redim I1_b_plant(P062_I1_s_o_flag)

If Request("txtPlantCd2") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)							'⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
'-----------------------
'Data manipulate area
'-----------------------
I1_b_plant(P062_I1_plant_cd)	= UCase(Trim(Request("txtPlantCd2")))	                '☆: Plant Code
I1_b_plant(P062_I1_plant_nm)	= Trim(Request("txtPlantNm2"))
I1_b_plant(P062_I1_cur_cd)		= UCase(Trim(Request("txtCurCd")))
I1_b_plant(P062_I1_plan_hrzn)	= UniConvNum(Request("txtPlngHrzn"), 0)
I1_b_plant(P062_I1_ptf_for_mps)	= UniCInt(Request("txtPtfForMps"), 0)
I1_b_plant(P062_I1_dtf_for_mps)	= UniCInt(Request("txtDtfForMps"), 0) 
I1_b_plant(P062_I1_ptf_for_mrp)	= UniCInt(Request("txtPtfForMrp"), 0)

I1_b_plant(P062_I1_inv_open_dt)	= UniConvDate(Request("hInvOpenDt"))
I1_b_plant(P062_I1_inv_cls_dt)	= UniConvDate(Request("hInvClsDt"))
I1_b_plant(P062_I1_country_cd)	= UCase(Trim(Request("txtCountryCd")))
I1_b_plant(P062_I1_s_o_flag) = Trim(Request("cboSOFlag"))

If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I1_b_plant(P062_I1_valid_from_dt) = UniConvDate(Request("txtValidFromDt"))
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I1_b_plant(P062_I1_valid_to_dt)   = UniConvDate(Request("txtValidToDt"))
	End If
End If

I2_biz_area_cd	= UCase(Trim(Request("txtBizAreaCd")))

I3_cal_type	= UCase(Trim(Request("txtClnrType")))

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Set pPB6G102 = Server.CreateObject("PB6G102.cBMngPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB6G102.B_MANAGE_PLANT(gStrGlobalCollection, iCommandSent, I1_b_plant, I2_biz_area_cd, I3_cal_type)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB6G102 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB6G102 = Nothing															'☜: Unload Component
	
'-----------------------
'Result data display area
'----------------------- 
Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>
