<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b03mb2.asp	
'*  4. Program Name         : Entry ItemGroup (Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PB3G103.cMngItemGrp
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear

Dim pPB3G103																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_b_item_group, I2_upper_item_group_cd, iCommandSent
Dim iIntFlgMode

Const P019_I1_item_group_cd = 0
Const P019_I1_item_group_nm = 1
Const P019_I1_leaf_flg = 2
Const P019_I1_valid_from_dt = 3
Const P019_I1_valid_to_dt = 4

If Request("txtItemGroupCd2") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End
End If

Redim I1_b_item_group(P019_I1_valid_to_dt)
    
iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

'-----------------------
'Data manipulate area
'-----------------------
I1_b_item_group(P019_I1_item_group_cd) = UCase(Trim(Request("txtItemGroupCd2")))
I1_b_item_group(P019_I1_item_group_nm) = Trim(Request("txtItemGroupNm2"))
I1_b_item_group(P019_I1_leaf_flg) = Trim(Request("rdoLowItemGroupFlg"))

If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I1_b_item_group(P019_I1_valid_from_dt) = UniConvDate(Request("txtValidFromDt"))
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I1_b_item_group(P019_I1_valid_to_dt) = UniConvDate(Request("txtValidToDt"))
	End If
End If

I2_upper_item_group_cd	= Trim(UCase(Request("txtHighItemGroupCd")))

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If
    
'-----------------------
'Com Action Area
'-----------------------
Set pPB3G103 = Server.CreateObject("PB3G103.cBMngItemGrp")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3G103.B_MANAGE_ITEM_GROUP(gStrGlobalCollection, iCommandSent, I1_b_item_group, I2_upper_item_group_cd)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3G103 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB3G103 = Nothing															'☜: Unload Component

'-----------------------
'Result data display area
'----------------------- 
Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>
