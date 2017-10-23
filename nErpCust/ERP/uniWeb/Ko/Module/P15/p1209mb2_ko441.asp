<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1204mb2.asp
'*  4. Program Name         : Entry Standard RoutingEntry (Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PP1G507.cPMngStdRtng
'*  7. Modified date(First) : 2000/04/3
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1G507																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_rout_no, I3_major_flg, I4_cost_center, I5_altrtvalue
Dim iErrorPosition, strSpread

If Request("txtPlantCd") = "" Then													'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 
ElseIf Request("txtMaxRows") = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)                     
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------
strSpread   = Request("txtSpread")
I1_plant_cd = UCase(Trim(Request("txtPlantCd")))
I2_rout_no = UCase(Trim(Request("txtRoutingNo")))
	
If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	End If
End If

I3_major_flg	= UCase(Trim(Request("rdoMajorRouting")))
I4_cost_center	= UCase(Trim(Request("txtCostCd")))
I5_altrtvalue	= UNIConvNum(Request("txtALTRTVALUE"), 0)

Set pPP1G507 = Server.CreateObject("PP1G507_KO441.cPMngItemGrpStdRtng")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1G507.P_MANAGE_STANDARD_ROUTING(gStrGlobalCollection, "", strSpread, _
                               I1_plant_cd, I2_rout_no, I3_major_flg, I4_cost_center, I5_altrtvalue, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPP1G507 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1G507 = Nothing																'☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>