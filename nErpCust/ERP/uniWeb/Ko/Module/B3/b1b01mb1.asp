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
'*  3. Program ID           : b1b01mb1.asp
'*  4. Program Name         : Entry Item(Query)
'*  5. Program Desc         :
'*  6. Component List       : PB3C104.cBLkUpItem
'*  7. Modified date(First) : 2000/03/25
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3C104																'☆ : 조회용 Component Dll 사용 변수 

Dim I1_select_char, I2_item_cd
Dim E1_b_minor, E2_select_char, E3_base_b_item, E4_b_item_group, E5_b_item, iStatusCodeOfPrevNext

' E1_b_minor
Const P022_E1_minor_cd = 0
Const P022_E1_minor_nm = 1

' E3_base_b_item
Const P022_E3_item_cd = 0
Const P022_E3_item_nm = 1

' E4_b_item_group
Const P022_E4_item_group_cd = 0
Const P022_E4_item_group_nm = 1
Const P022_E4_leaf_flg = 2
Const P022_E4_del_flg = 3
Const P022_E4_valid_from_dt = 4
Const P022_E4_valid_to_dt = 5

' E5_b_item
Const P022_E5_item_cd = 0
Const P022_E5_item_nm = 1
Const P022_E5_formal_nm = 2
Const P022_E5_spec = 3
Const P022_E5_basic_unit = 4
Const P022_E5_item_acct = 5
Const P022_E5_item_class = 6
Const P022_E5_phantom_flg = 7
Const P022_E5_hs_cd = 8
Const P022_E5_hs_unit = 9
Const P022_E5_unit_weight = 10
Const P022_E5_unit_of_weight = 11
Const P022_E5_draw_no = 12
Const P022_E5_item_image_flg = 13
Const P022_E5_blanket_pur_flg = 14
Const P022_E5_base_item_cd = 15
Const P022_E5_proportion_rate = 16
Const P022_E5_valid_flg = 17
Const P022_E5_valid_from_dt = 18
Const P022_E5_valid_to_dt = 19
Const P022_E5_vat_type = 20
Const P022_E5_vat_rate = 21
Const P022_E5_unit_gross_weight = 22
Const P022_E5_unit_of_gross_weight = 23
Const P022_E5_cbm_volume = 24
Const P022_E5_cbm_info = 25

If Request("txtItemCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

'-----------------------
'Data manipulate  area(import view match)
'-----------------------
    
I1_select_char = Request("PrevNextFlg")
I2_item_cd = Trim(Request("txtItemCd"))

'-----------------------
'Com action area
'-----------------------
Set pPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3C104.B_LOOK_UP_ITEM_SVR(gStrGlobalCollection, I1_select_char, I2_item_cd, E1_b_minor, _
                     E2_select_char, E3_base_b_item, E4_b_item_group, E5_b_item, iStatusCodeOfPrevNext)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3C104 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB3C104 = Nothing															'☜: Unload Component

If Trim(iStatusCodeOfPrevNext) = "900011" Or Trim(iStatusCodeOfPrevNext) = "900012" Then
	Call DisplayMsgBox(iStatusCodeOfPrevNext, VbOKOnly, "", "", I_MKSCRIPT)
End If
	
'-----------------------
'Result data display area
'----------------------- 
' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.
Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtItemCd.Value = """ & ConvSPChars(E5_b_item(P022_E5_item_cd)) & """" & vbCrLf		
		Response.Write ".txtItemNm.Value = """ & ConvSPChars(E5_b_item(P022_E5_item_nm)) & """" & vbCrLf
		Response.Write ".txtItemCd1.Value = """ & ConvSPChars(E5_b_item(P022_E5_item_cd)) & """" & vbCrLf
		Response.Write ".txtItemNm1.Value = """ & ConvSPChars(E5_b_item(P022_E5_item_nm)) & """" & vbCrLf
		Response.Write ".txtItemDesc.Value = """ & ConvSPChars(E5_b_item(P022_E5_formal_nm)) & """" & vbCrLf
		Response.Write ".txtUnit.Value = """ & ConvSPChars(E5_b_item(P022_E5_basic_unit)) & """" & vbCrLf
		Response.Write ".cboItemAcct.Value = """ & E5_b_item(P022_E5_item_acct) & """" & vbCrLf
		Response.Write ".txtItemGroupCd.Value = """ & ConvSPChars(E4_b_item_group(P022_E4_item_group_cd)) & """" & vbCrLf
		Response.Write ".txtItemGroupNm.Value = """ & ConvSPChars(E4_b_item_group(P022_E4_item_group_nm)) & """" & vbCrLf
		Response.Write ".txtItemByPlantFlg.Value = """ & E2_select_char & """" & vbCrLf

		If E5_b_item(P022_E5_phantom_flg) = "Y" Then
			Response.Write ".rdoPhantomType(0).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal1 = 1" & vbCrLf
		Else
			Response.Write ".rdoPhantomType(1).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal1 = 2" & vbCrLf
		End If		

		If E5_b_item(P022_E5_blanket_pur_flg) = "Y" Then
			Response.Write ".rdoUnifyPurFlg(0).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal2 = 1" & vbCrLf
		Else
			Response.Write ".rdoUnifyPurFlg(1).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal2 = 2" & vbCrLf
		End If
 		
		Response.Write ".txtBasisItemCd.Value = """ & ConvSPChars(E5_b_item(P022_E5_base_item_cd)) & """" & vbCrLf
		Response.Write ".txtBasisItemNm.Value = """ & ConvSPChars(E3_base_b_item(P022_E3_item_nm)) & """" & vbCrLf
		Response.Write ".cboItemClass.Value = """ & E5_b_item(P022_E5_item_class) & """" & vbCrLf
		Response.Write ".txtValidFromDt.Text = """ & UniDateClientFormat(E5_b_item(P022_E5_valid_from_dt)) & """" & vbCrLf
		Response.Write ".txtValidToDt.Text = """ & UniDateClientFormat(E5_b_item(P022_E5_valid_to_dt)) & """" & vbCrLf
		Response.Write ".txtItemSpec.Value = """ & ConvSPChars(E5_b_item(P022_E5_spec)) & """" & vbCrLf
		Response.Write ".txtWeight.Text = """ & UniConvNumberDBToCompany(E5_b_item(P022_E5_unit_weight), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtWeightUnit.Value = """ & ConvSPChars(E5_b_item(P022_E5_unit_of_weight)) & """" & vbCrLf
		Response.Write ".txtGrossWeight.Text = """ & UniConvNumberDBToCompany(E5_b_item(P022_E5_unit_gross_weight), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtGrossWeightUnit.Value = """ & ConvSPChars(E5_b_item(P022_E5_unit_of_gross_weight)) & """" & vbCrLf
		Response.Write ".txtCBM.Text = """ & UniConvNumberDBToCompany(E5_b_item(P022_E5_cbm_volume), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtCBMInfo.Value = """ & ConvSPChars(E5_b_item(P022_E5_cbm_info)) & """" & vbCrLf
		Response.Write ".txtDrawNo.Value = """ & ConvSPChars(E5_b_item(P022_E5_draw_no)) & """" & vbCrLf
		If E5_b_item(P022_E5_valid_flg) = "Y" Then
			Response.Write ".rdoValidFlg(0).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal3 = 1" & vbCrLf
		Else
			Response.Write ".rdoValidFlg(1).Checked = True" & vbCrLf
			Response.Write "parent.lgRdoOldVal3 = 2" & vbCrLf
		End If
		
		Response.Write ".txtHsCd.Value = """ & ConvSPChars(E5_b_item(P022_E5_hs_cd)) & """" & vbCrLf
		Response.Write ".txtHsUnit.Value = """ & ConvSPChars(E5_b_item(P022_E5_hs_unit)) & """" & vbCrLf

		If E5_b_item(P022_E5_item_image_flg) = "Y" Then
			Response.Write ".rdoPhoto(0).Checked = True" & vbCrLf
		Else
			Response.Write ".rdoPhoto(1).Checked = True" & vbCrLf
		End If
		
		Response.Write ".txtVatType.Value = """ & ConvSPChars(E5_b_item(P022_E5_vat_type)) & """" & vbCrLf
		Response.Write ".txtVatTypeNm.Value = """ & ConvSPChars(E1_b_minor(P022_E1_minor_nm)) & """" & vbCrLf
		Response.Write ".txtVatRate.Text = """ & UniConvNumberDBToCompany(E5_b_item(P022_E5_vat_rate), ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0) & """" & vbCrLf
		
		Response.Write "parent.DbQueryOk" & vbCrLf	'☜: 조회가 성공 
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End										'☜: Process End
%>