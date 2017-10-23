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
'*  3. Program ID           : p1301mb2.asp	
'*  4. Program Name         : Entry Work Center (Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PP1G203.cMngWkCtr
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPP1G203																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_cost_cd, I3_cal_type, I4_p_work_center, iCommandSent
Dim iIntFlgMode

Const P118_I4_wc_cd = 0
Const P118_I4_wc_nm = 1
Const P118_I4_inside_flg = 2
Const P118_I4_wc_mgr = 3
Const P118_I4_valid_from_dt = 4
Const P118_I4_valid_to_dt = 5

Redim I4_p_work_center(P118_I4_valid_to_dt)

If Request("txtFlgMode") = "" Then														'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)  '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
If Request("txtDataWcCd") = "" Then														'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)		'⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
If Request("txtClnrType") = "" Then														'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)		'⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
iIntFlgMode = CInt(Request("txtFlgMode"))												'☜: 저장시 Create/Update 판별 

'-----------------------
'Data manipulate area
'-----------------------
I1_plant_cd = UCase(Trim(Request("txtPlantCd")))
I2_cost_cd = UCase(Trim(Request("txtCostCd")))
I3_cal_type = UCase(Trim(Request("txtClnrType")))

I4_p_work_center(P118_I4_wc_cd) = UCase(Trim(Request("txtDataWcCd")))
I4_p_work_center(P118_I4_wc_nm) = Trim(Request("txtDataWcNm"))
I4_p_work_center(P118_I4_wc_mgr) = Trim(Request("cboWcMgr"))
I4_p_work_center(P118_I4_inside_flg) = UCase(Trim(Request("cboInsideFlg")))
	
If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I4_p_work_center(P118_I4_valid_from_dt) = UNIConvDate(Request("txtValidFromDt"))
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I4_p_work_center(P118_I4_valid_to_dt) = UNIConvDate(Request("txtValidToDt"))
	End If
End If
	
If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Set pPP1G203 = Server.CreateObject("PP1G203.cPMngWkCtr")

If CheckSYSTEMError(Err, True) = True Then
	Response.End
End If

Call pPP1G203.P_MANAGE_WORK_CENTER(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_cost_cd, _
                                   I3_cal_type, I4_p_work_center)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1G203 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1G203 = Nothing																'☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>