<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b03mb1.asp
'*  4. Program Name         : Entry ItemGroup (Query)
'*  5. Program Desc         :
'*  6. Component List       : PB3S101.cBLkUpItemGrp
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2003/01/06
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3S101																'☆ : 조회용 Component Dll 사용 변수 

Dim I1_select_char, I2_item_group_cd
Dim E1_b_item_group, E2_b_item_group, iStatusCodeOfPrevNext

Const P014_E1_item_group_cd = 0
Const P014_E1_item_group_nm = 1
Const P014_E1_leaf_flg = 2
Const P014_E1_valid_from_dt = 3
Const P014_E1_valid_to_dt = 4
Const P014_E1_level = 5

Const P014_E2_item_group_cd = 0
Const P014_E2_item_group_nm = 1

If Request("txtItemGroupCd1") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)       
	Response.End 
End If
        
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_select_char = Request("PrevNextFlg")
I2_item_group_cd = UCase(Trim(Request("txtItemGroupCd1")))

'-----------------------
'Com action area
'-----------------------
Set pPB3S101 = Server.CreateObject("PB3S101.cBLkUpItemGrp")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S101.B_LOOK_UP_ITEM_GROUP_SVR(gStrGlobalCollection, I1_select_char, I2_item_group_cd, _
                     E1_b_item_group, E2_b_item_group, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S101 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB3S101 = Nothing															'☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If

Response.Write "<Script Language=VBScript>" & vbCRLF
Response.Write "With parent.frm1" & vbCRLF
	Response.Write ".txtItemGroupCd1.value = """ & ConvSPChars(E1_b_item_group(P014_E1_item_group_cd)) & """" & vbCRLF
	Response.Write ".txtItemGroupNm1.value = """ & ConvSPChars(E1_b_item_group(P014_E1_item_group_nm)) & """" & vbCRLF
	Response.Write ".txtItemGroupCd2.value = """ & ConvSPChars(E1_b_item_group(P014_E1_item_group_cd)) & """" & vbCRLF
	Response.Write ".txtItemGroupNm2.value = """ & ConvSPChars(E1_b_item_group(P014_E1_item_group_nm)) & """" & vbCRLF
	Response.Write ".txtHighItemGroupCd.value = """ & ConvSPChars(E2_b_item_group(P014_E2_item_group_cd)) & """" & vbCRLF
	Response.Write ".txtHighItemGroupNm.value = """ & ConvSPChars(E2_b_item_group(P014_E2_item_group_nm)) & """" & vbCRLF
	If Trim(E1_b_item_group(P014_E1_leaf_flg)) = "Y" Then
		Response.Write ".rdoLowItemGroupFlg1.checked = True" & vbCRLF
		Response.Write "parent.lgRdoOldVal1 = 1" & vbCRLF
	Else
		Response.Write ".rdoLowItemGroupFlg2.checked = True" & vbCRLF
		Response.Write "parent.lgRdoOldVal1 = 2" & vbCRLF
	End If
	Response.Write ".txtValidFromDt.text = """ & UniDateClientFormat(E1_b_item_group(P014_E1_valid_from_dt)) & """" & vbCRLF
	Response.Write ".txtValidToDt.text = """ & UniDateClientFormat(E1_b_item_group(P014_E1_valid_to_dt)) & """" & vbCRLF
	Response.Write ".txtlevel1.value = """ & E1_b_item_group(P014_E1_level) & """" & vbCRLF
	Response.Write "parent.DbQueryOk" & vbCRLF
Response.Write "End With" & vbCRLF
Response.Write "</Script>" & vbCRLF
Response.End																	'☜: Process End
%>