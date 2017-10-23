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
'*  3. Program ID           : b1b11mb9_ko441.asp	
'*  4. Program Name         : Entry Item By Routing(Create)
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

Dim pPP1C509																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_item_cd, I3_item_grp1, I4_item_grp2, i5_user_id, iCommandSent
Dim iIntFlgMode, itxtMode

'-------------------------------------------------------------------------
' Validation Check
'-------------------------------------------------------------------------
If Request("txtPlantCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT) '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If

If Request("txtItemCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT) '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If

itxtMode	= Request("txtMode")										'☜: 저장시 Create/Update 판별 

I1_plant_cd	= UCase(Trim(Request("txtPlantCd")))
I2_item_cd	= UCase(Trim(Request("txtItemCd")))
I3_item_grp1	= UCase(Trim(Request("txtItemGrp1")))
I4_item_grp2	= UCase(Trim(Request("txtItemGrp2")))
i5_user_id	= UCase(Trim(Request("txtInsrtUserId")))

If itxtMode = "CrtRouting" Then																	 
	iCommandSent = "CREATE"						
Else
	iCommandSent = "UPDATE"
	Response.End
End If

Set pPP1C509 = Server.CreateObject("PP1S509_KO441.cPCrtAutoRouting")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1C509.P_CREATE_ITEM_ROUTING(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_item_cd, I3_item_grp1, I4_item_grp2, i5_user_id)

'response.write " b1b11mb9_ko441.asp iIntFlgMode->" & iIntFlgMode


If CheckSYSTEMError(Err, True) = True Then
	Set pPP1C509 = Nothing
	Response.End
End If

Set pPP1C509 = Nothing


Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.RunCrtRoutingOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>
