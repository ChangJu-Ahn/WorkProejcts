<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : b1b03mb4.asp
'*  4. Program Name         :
'*  5. Program Desc         : 품목그룹조회 
'*  6. Component List       : PB3G112.cBListItemGrp.B_LIST_ITEM_GROUP
'*  7. Modified date(First) : 2000/04/28
'*  8. Modified date(Last)  : 2000/09/28
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

Response.Flush 

'[CONVERSION INFORMATION]  EXPORTS View 상수 
    
'[CONVERSION INFORMATION]  View Name : export b_item_group
Const P039_E1_item_group_cd = 0
Const P039_E1_item_group_nm = 1
Const P039_E1_leaf_flg = 2
Const P039_E1_valid_from_dt = 3
Const P039_E1_valid_to_dt = 4
Const P039_E1_level = 5
    
'[CONVERSION INFORMATION]  View Name : export_upper b_item_group
Const P039_E2_item_group_cd = 0
Const P039_E2_item_group_nm = 1
Const P039_E2_leaf_flg = 2
Const P039_E2_valid_from_dt = 3
Const P039_E2_valid_to_dt = 4
Const P039_E2_level = 5
    
'[CONVERSION INFORMATION]  View Name : export_next p_bom_for_explosion
Const P039_E4_seq = 0

'[CONVERSION INFORMATION]  EXPORTS Group View 상수 
'[CONVERSION INFORMATION]  Group Name : export_group
'[CONVERSION INFORMATION]  View Name : export_item p_bom_for_explosion
Const P039_EG1_E1_p_bom_for_explosion_seq = 0
Const P039_EG1_E1_p_bom_for_explosion_user_id = 1
Const P039_EG1_E1_p_bom_for_explosion_plant_cd = 2
Const P039_EG1_E1_p_bom_for_explosion_prnt_item_cd = 3
Const P039_EG1_E1_p_bom_for_explosion_prnt_bom_no = 4
Const P039_EG1_E1_p_bom_for_explosion_level_cd = 5
Const P039_EG1_E1_p_bom_for_explosion_child_item_cd = 6
Const P039_EG1_E1_p_bom_for_explosion_prnt_node = 7
Const P039_EG1_E1_p_bom_for_explosion_own_node = 8
Const P039_EG1_E1_p_bom_for_explosion_material_flg = 9
Const P039_EG1_E1_p_bom_for_explosion_bom_flg = 10

Call LoadBasisGlobalInf() 
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                     '☜: Hide Processing message

On Error Resume Next														'☜: 

Dim pPB3G112																'☆ : 입력/수정용 ComProxy Dll 사용 변수 

Dim LngRow
Dim iIntCnt, strTemp, iIntLevel

Dim I1_srch_type_select_char
Dim I2_b_plant_plant_cd 
Dim I3_p_bom_for_explosion_seq 
Dim I4_next_flag_select_char 
Dim I5_b_item_group_cd 
Dim E1_b_item_group
Dim E2_b_item_group
Dim EG1_export_group
Dim iStatusCodeOfPrevNext
Dim iErrorPosition

If CInt(Request("txtMode")) <> UID_M0001 Then
	Response.End 
End If

I2_b_plant_plant_cd = Request("txtPlantCd")										' 조회할 키 
I5_b_item_group_cd = Request("txtItemGroupCd")									' 조회할 상위키 
	
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I3_p_bom_for_explosion_seq  = 0
I1_srch_type_select_char = Request("txtSrchType")
		
'-----------------------
'Com Action Area
'-----------------------
		
Set pPB3G112=Server.CreateObject ("PB3G112.cBListItemGrp")
If CheckSYSTEMError(Err, True) = True Then
	Response.End
End if
	
Call pPB3G112.B_LIST_ITEM_GROUP (gStrGlobalCollection, Cstr(I1_srch_type_select_char), _
		Cstr(I2_b_plant_plant_cd), Cstr(I3_p_bom_for_explosion_seq), Cstr(I4_next_flag_select_char), _
		Cstr(I5_b_item_group_cd), E1_b_item_group, E2_b_item_group, EG1_export_group, iStatusCodeOfPrevNext)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPB3G112 = Nothing	
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If
	
Set pPB3G112 = Nothing												'☜: Unload Component

'------------------------------
' Level Setting
'------------------------------
If Trim(E2_b_item_group(P039_E2_level)) <> "0" Then
	For iIntCnt = 1 To CInt(E2_b_item_group(P039_E2_level))
		strTemp = strTemp & "."			
	Next

	iIntLevel = strTemp & CStr(E2_b_item_group(P039_E2_level))
Else
	iIntLevel = CStr(E2_b_item_group(P039_E2_level))
End If 
		
'--------------------------------
'Next가 아니면 Header정보 Setting
'--------------------------------
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtLevel1.value = """ & iIntLevel & """" & vbCrLf
		Response.Write ".hRootLevel.value = """ & Trim(E2_b_item_group(P039_E2_level)) & """" & vbCrLf
		Response.Write ".txtItemGroupNm.value =	""" & ConvSPChars(E2_b_item_group(P039_E2_item_group_nm)) & """" & vbCrLf
		Response.Write ".txtItemGroupCd1.value = """ & ConvSPChars(E2_b_item_group(P039_E2_item_group_cd)) & """" & vbCrLf
		Response.Write ".txtItemGroupNm1.value = """ & ConvSPChars(E2_b_item_group(P039_E2_item_group_nm)) & """" & vbCrLf
		Response.Write ".txtUpperItemGroupCd.value = """ & ConvSPChars(E1_b_item_group(P039_E1_item_group_cd)) & """" & vbCrLf
		Response.Write ".txtUpperItemGroupNm.value = """ & ConvSPChars(E1_b_item_group(P039_E1_item_group_nm)) & """" & vbCrLf
		If	UCase(Trim(E1_b_item_group(P039_E1_leaf_flg)))  = "Y" Then 
			Response.Write ".rdoLowItemGroupFlg1.Checked = True" & vbCrLf
		Else
			Response.Write ".rdoLowItemGroupFlg2.Checked = True" & vbCrLf
		End If
		Response.Write ".txtValidFromDt1.value = """ & UNIDateClientFormat(E2_b_item_group(P039_E2_valid_from_dt)) & """" & vbCrLf
		Response.Write ".txtValidToDt1.value = """ & UNIDateClientFormat(E2_b_item_group(P039_E2_valid_to_dt)) & """" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

'----------------------------------------------------
'- Parent Node를 Setting하고 Header Data를 가져온다.
'---------------------------------------------------
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "Dim PrntKey" & vbCrLf
	Response.Write "Dim NodX" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
				      
		Response.Write "PrntKey = ""0|^|^|0""" & vbCrLf

		Response.Write "Set NodX = .uniTree1.Nodes.Add(,, PrntKey, """ & ConvSPChars(UCase(Trim(I5_b_item_group_cd))) & """, parent.C_GROUP, parent.C_GROUP)" & vbCrLf
		Response.Write "NodX.Expanded = True" & vbCrLf

		Response.Write "Set NodX = Nothing" & vbCrLf

	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

If Not IsEmpty(EG1_export_group) Then
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "Dim Node" & vbCrLf
	Response.Write "With parent.frm1.uniTree1" & vbCrLf
	Response.Write ".MousePointer = 11" & vbCrLf													'⊙: 마우스 포인트 변화 
	Response.Write ".Indentation = 50" & vbCrLf		
	For LngRow = 0 to ubound(EG1_export_group, 1)
		If EG1_export_group(LngRow, P039_EG1_E1_p_bom_for_explosion_bom_flg) = "0" Then		' 제품일 경우 
			Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_prnt_node))) & """, _ " & vbCrLf
			Response.Write "parent.tvwChild,""" & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_own_node))) & """, _ " & vbCrLf
			Response.Write """" & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_child_item_cd))) & """, parent.C_GROUP, parent.C_GROUP)" & vbCrLf
			Response.Write "Node.Expanded = True " & vbCrLf
		Else																' 원자재인 경우 
			Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_prnt_node))) & """, _ " & vbCrLf
			Response.Write "parent.tvwChild, """ & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_own_node))) & """, _ " & vbCrLf
			Response.Write """" & ConvSPChars(Trim(EG1_export_group(LngRow,P039_EG1_E1_p_bom_for_explosion_child_item_cd))) & """, parent.C_PROD, parent.C_PROD)" & vbCrLf
			Response.Write "Node.Expanded = True" & vbCrLf
		End If
	Next
	Response.Write ".MousePointer = 1" & vbCrLf
	Response.Write "Set Node = Nothing" & vbCrLf

	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "Call parent.DbQueryOk()" & vbCrLf
Response.Write "</Script>" & vbCrLf
%>
