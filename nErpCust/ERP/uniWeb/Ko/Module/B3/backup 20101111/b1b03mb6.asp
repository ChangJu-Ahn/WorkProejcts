<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01mb1.asp
'*  4. Program Name         : Item Query
'*  5. Program Desc         :
'*  6. Component List        : PB3C104.cBLkUpItem.B_LOOK_UP_ITEM_SVR
'*  7. Modified date(First) : 2000/03/25
'*  8. Modified date(Last)  : 2003/01/06
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

' E1_b_minor
Const P020_E1_minor_cd = 0
Const P020_E1_minor_nm = 1

' E3_base_b_item
Const P020_E3_item_cd = 0
Const P020_E3_item_nm = 1

' E4_b_item_group
Const P020_E4_item_group_cd = 0
Const P020_E4_item_group_nm = 1
Const P020_E4_leaf_flg = 2
Const P020_E4_del_flg = 3
Const P020_E4_valid_from_dt = 4
Const P020_E4_valid_to_dt = 5

'E5_b_item
Const P020_E5_item_cd = 0
Const P020_E5_item_nm = 1
Const P020_E5_formal_nm = 2
Const P020_E5_spec = 3
Const P020_E5_basic_unit = 4
Const P020_E5_item_acct = 5
Const P020_E5_item_class = 6
Const P020_E5_phantom_flg = 7
Const P020_E5_hs_cd = 8
Const P020_E5_hs_unit = 9
Const P020_E5_unit_weight = 10
Const P020_E5_unit_of_weight = 11
Const P020_E5_draw_no = 12
Const P020_E5_item_image_flg = 13
Const P020_E5_blanket_pur_flg = 14
Const P020_E5_base_item_cd = 15
Const P020_E5_proportion_rate = 16
Const P020_E5_valid_flg = 17
Const P020_E5_valid_from_dt = 18
Const P020_E5_valid_to_dt = 19
Const P020_E5_vat_type = 20
Const P020_E5_vat_rate = 21
Const P020_E5_gross_weight = 22
Const P020_E5_unit_of_gross_weight = 23
Const P020_E5_cbm_volume = 24
Const P020_E5_cbm_info = 25

On Error Resume Next														'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Call LoadBasisGlobalInf() 
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "MB")

Dim pPB3C104																'☆ : 조회용 ComProxy Dll 사용 변수 
Dim I1_b_item_cd
Dim E1_b_minor
Dim E2_select_char
Dim E3_base_b_item
Dim E4_b_item_group
Dim E5_b_item
Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim i
Dim strTemp
Dim intLevel
Dim iLevelCnt

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Err.Clear                                                               '☜: Protect system from crashing
    
If Request("txtItemCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)       
	Response.End 
End If

'------------------------------
' Level Setting
'------------------------------
iLevelCnt = CInt(Request("txtLevelCd")) + CInt(Request("txtRootLevel"))
If iLevelCnt > 0 Then
	For i = 1 To iLevelCnt
		strTemp = strTemp & "."			
	Next
	
	intLevel = strTemp & CStr(iLevelCnt)
End If    
	
I1_b_item_cd		= Trim(Request("txtItemCd"))
    
Set pPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")       

'-----------------------
'Com action result check area(OS,internal)    
'-----------------------
     
If CheckSYSTEMError(Err, True) = True Then
	Response.End 
End if

'-----------------------
'Com action area
'-----------------------
Call pPB3C104.B_LOOK_UP_ITEM  (gStrGlobalCollection, I1_b_item_cd, E1_b_minor, E2_select_char, _
        E3_base_b_item, E4_b_item_group, E5_b_item)
   
'-----------------------
'Com action result check area(OS,internal)	
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Set pPB3C104 = Nothing												'☜: ComProxy Unload
	Response.End
End If
    
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtLevel2.value = """ & intLevel & """" & vbCrLf
		Response.Write ".txtItemCd.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_item_cd))) & """" & vbCrLf
		Response.Write ".txtItemNm.value = """ & ConvSPChars(E5_b_item(P020_E5_item_nm)) & """" & vbCrLf
		Response.Write ".txtItemDesc.value = """ & ConvSPChars(E5_b_item(P020_E5_formal_nm)) & """" & vbCrLf
		Response.Write ".txtBasicUnit.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_basic_unit))) & """" & vbCrLf
		Response.Write ".cboItemAcct.value = """ & UCase(E5_b_item(P020_E5_item_acct)) & """" & vbCrLf
		
		Response.Write ".txtItemGroupCd2.value = """ & ConvSPChars(E4_b_item_group(P020_E4_item_group_cd)) & """" & vbCrLf
		Response.Write ".txtItemGroupNm2.value = """ & ConvSPChars(E4_b_item_group(P020_E4_item_group_nm)) & """" & vbCrLf

		If UCase(E5_b_item(P020_E5_phantom_flg)) = "Y" Then
			Response.Write ".rdoPhantomType(0).checked = True" & vbCrLf
		Else
			Response.Write ".rdoPhantomType(1).checked = True" & vbCrLf
		End If		

		If UCase(E5_b_item(P020_E5_phantom_flg)) = "Y" Then
			Response.Write ".rdoUnifyPurFlg(0).checked = True" & vbCrLf
		Else
			Response.Write ".rdoUnifyPurFlg(1).checked = True" & vbCrLf
		End If
 		 		
		Response.Write ".txtBasisItemCd.value = """ & ConvSPChars(UCase(E3_base_b_item(P020_E3_item_cd))) & """" & vbCrLf
		Response.Write ".txtBasisItemNm.value = """ & ConvSPChars(E3_base_b_item(P020_E3_item_nm)) & """" & vbCrLf
		Response.Write ".cboItemClass.value = """ & UCase(E5_b_item(P020_E5_item_class)) & """" & vbCrLf
		Response.Write ".txtItemSpec.value = """ & ConvSPChars(E5_b_item(P020_E5_spec)) & """" & vbCrLf
		Response.Write ".txtWeight.value = """ & UniConvNumberDBToCompany(E5_b_item(P020_E5_unit_weight), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtWeightUnit.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_unit_of_weight))) & """" & vbCrLf
		Response.Write ".txtGrossWeight.value = """ & UniConvNumberDBToCompany(E5_b_item(P020_E5_gross_weight), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtGrossWeightUnit.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_unit_of_gross_weight))) & """" & vbCrLf
		Response.Write ".txtCBM.value = """ & UniConvNumberDBToCompany(E5_b_item(P020_E5_cbm_volume), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & """" & vbCrLf
		Response.Write ".txtCBMInfo.value = """ & ConvSPChars(E5_b_item(P020_E5_cbm_info)) & """" & vbCrLf
		Response.Write ".txtDrawNo.value = """ & ConvSPChars(E5_b_item(P020_E5_draw_no)) & """" & vbCrLf

		If UCase(E5_b_item(P020_E5_valid_flg)) = "Y" Then
			Response.Write ".rdoValidFlg(0).checked = True" & vbCrLf
		Else
			Response.Write ".rdoValidFlg(1).checked = True" & vbCrLf
		End If
		
		Response.Write ".txtHsCd.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_hs_cd))) & """" & vbCrLf
		Response.Write ".txtHsUnit.value = """ & ConvSPChars(UCase(E5_b_item(P020_E5_hs_unit))) & """" & vbCrLf
		If UCase(E5_b_item(P020_E5_item_image_flg)) = "Y" Then
			Response.Write ".rdoPhoto(0).checked = True" & vbCrLf
		Else
			Response.Write ".rdoPhoto(1).checked = True" & vbCrLf
		End If

		Response.Write ".txtValidFromDt2.value = """ & UniDateClientFormat(E5_b_item(P020_E5_valid_from_dt)) & """" & vbCrLf
		Response.Write ".txtValidToDt2.value = """ & UniDateClientFormat(E5_b_item(P020_E5_valid_to_dt)) & """" & vbCrLf
		
		Response.Write "parent.LookUpItemOk" & vbCrLf														'☜: 조화가 성공 
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																	'☜: Process End
%>
