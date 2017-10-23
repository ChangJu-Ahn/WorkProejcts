<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1104mb2.asp
'*  4. Program Name         : Entry Shift(Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PP1G602.cPMngShift
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1G602																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_p_shift_header, iCommandSent, iErrorPosition, strSpread
Dim strPlantCd, iIntFlgMode

Const P159_I2_shift_cd = 0
Const P159_I2_description = 1
Const P159_I2_valid_from_dt = 2
Const P159_I2_valid_to_dt = 3

Redim I2_p_shift_header(P159_I2_valid_to_dt)

iIntFlgMode = CInt(Request("txtFlgMode"))
    
If iIntFlgMode = OPMD_CMODE Then
	strPlantCd = Request("txtPlantCd")
Else
	strPlantCd = Request("hPlantCd")
End If

If strPlantCd = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If

If Request("txtShiftCd2") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
End If

'Data manipulate area
strSpread = Request("txtSpread")
'Shift Header Data 
I1_plant_cd	= UCase(Trim(strPlantCd))
I2_p_shift_header(P159_I2_shift_cd)			= UCase(Trim(Request("txtShiftCd2")))
I2_p_shift_header(P159_I2_description)		= Request("txtShiftNm2")
I2_p_shift_header(P159_I2_valid_from_dt)	= UniConvDate(Request("txtValidFromDt"))
I2_p_shift_header(P159_I2_valid_to_dt)		= UniConvDate(Request("txtValidToDt"))
	
If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Set pPP1G602 = Server.CreateObject("PP1G602.cPMngShift")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1G602.P_MANAGE_SHIFT(gStrGlobalCollection, iCommandSent, strSpread, I1_plant_cd, I2_p_shift_header, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPP1G602 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1G602 = Nothing																'☜: Unload Component

Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>