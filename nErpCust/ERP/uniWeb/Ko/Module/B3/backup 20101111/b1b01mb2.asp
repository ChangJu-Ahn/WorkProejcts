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
'*  3. Program ID           : b1b01mb2.asp
'*  4. Program Name         : Entry Item(Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PB3S105.cBMngItem
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3S105																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_item_group_cd, I2_b_item, iCommandSent
Dim iIntFlgMode

Const P025_I2_item_cd = 0
Const P025_I2_item_nm = 1
Const P025_I2_formal_nm = 2
Const P025_I2_phantom_flg = 3
Const P025_I2_item_acct = 4
Const P025_I2_item_class = 5
Const P025_I2_spec = 6
Const P025_I2_hs_cd = 7
Const P025_I2_hs_unit = 8
Const P025_I2_unit_weight = 9
Const P025_I2_unit_of_weight = 10
Const P025_I2_basic_unit = 11
Const P025_I2_draw_no = 12
Const P025_I2_blanket_pur_flg = 13
Const P025_I2_base_item_cd = 14
Const P025_I2_proportion_rate = 15
Const P025_I2_valid_flg = 16
Const P025_I2_valid_from_dt = 17
Const P025_I2_valid_to_dt = 18
Const P025_I2_vat_type = 19
Const P025_I2_vat_rate = 20
Const P025_I2_unit_gross_weight = 21
Const P025_I2_unit_of_gross_weight = 22
Const P025_I2_cbm_volume = 23
Const P025_I2_cbm_info = 24

If Request("txtItemCd1") = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)	 '⊙: 에러메세지는 DB화 한다.           
	Response.End 
End If
	
iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

'-----------------------
'Data manipulate area
'-----------------------
Redim I2_b_item(P025_I2_cbm_info)

I1_item_group_cd						= UCase(Trim(Request("txtItemGroupCd")))

I2_b_item(P025_I2_item_cd)				= UCase(Trim(Request("txtItemCd1")))
I2_b_item(P025_I2_item_nm)				= Trim(Request("txtItemNm1"))
I2_b_item(P025_I2_formal_nm)			= Trim(Request("txtItemDesc"))
I2_b_item(P025_I2_basic_unit)			= UCase(Trim(Request("txtUnit")))
I2_b_item(P025_I2_item_acct)			= UCase(Trim(Request("cboItemAcct")))
I2_b_item(P025_I2_base_item_cd)	        = UCase(Trim(Request("txtBasisItemCd")))
I2_b_item(P025_I2_item_class)			= Trim(Request("cboItemClass"))
I2_b_item(P025_I2_blanket_pur_flg)		= UCase(Request("rdoUnifyPurFlg"))
I2_b_item(P025_I2_proportion_rate)		= "0" 'Trim(Request("txtProportionRate"))
I2_b_item(P025_I2_valid_flg)			= UCase(Request("rdoValidFlg"))
    
If Trim(Request("cboItemAcct")) >= "30" Then
	I2_b_item(P025_I2_phantom_flg)	= "N" 	
Else
	I2_b_item(P025_I2_phantom_flg)	= UCase(Request("rdoPhantomType"))
End If
	
If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I2_b_item(P025_I2_valid_from_dt)	= UniConvDate(Request("txtValidFromDt"))
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I2_b_item(P025_I2_valid_to_dt)	= UniConvDate(Request("txtValidToDt"))  
	End If
End If

I2_b_item(P025_I2_spec)					= Trim(Request("txtItemSpec"))
I2_b_item(P025_I2_unit_weight)			= UniConvNum(Request("txtWeight"), 0)
I2_b_item(P025_I2_unit_of_weight) 		= UCase(Trim(Request("txtWeightUnit")))
I2_b_item(P025_I2_unit_gross_weight)	= UniConvNum(Request("txtGrossWeight"), 0)
I2_b_item(P025_I2_unit_of_gross_weight) = UCase(Trim(Request("txtGrossWeightUnit")))
I2_b_item(P025_I2_cbm_volume)			= UniConvNum(Request("txtCBM"), 0)
I2_b_item(P025_I2_cbm_info) 			= Trim(Request("txtCBMInfo"))
I2_b_item(P025_I2_draw_no)				= Trim(Request("txtDrawNo"))			                '☆: Plant Code
I2_b_item(P025_I2_hs_cd)				= Trim(Request("txtHsCd"))
I2_b_item(P025_I2_hs_unit)				= UCase(Trim(Request("txtHsUnit")))
I2_b_item(P025_I2_vat_type)				= UCase(Trim(Request("txtVatType")))
I2_b_item(P025_I2_vat_rate)				= UniConvNum(Request("txtVatRate"), 0)

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Set pPB3S105 = Server.CreateObject("PB3S105.cBMngItem")


If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3S105.B_MANAGE_ITEM(gStrGlobalCollection, iCommandSent, I1_item_group_cd, I2_b_item)

Select Case Trim(Cstr(Err.Description))
	Case "B_MESSAGE" & Chr(11) & "970023"
		Call DisplayMsgBox("970023", vbInformation, "품목종료일", "공장별품목종료일", I_MKSCRIPT)
		Set pPB3S105 = Nothing															'☜: Unload Component
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtValidToDt.focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	Case Else
		If CheckSYSTEMError(Err, True) = True Then
			Set pPB3S105 = Nothing															'☜: Unload Component
			Response.End
		End If
End Select

Set pPB3S105 = Nothing															'☜: Unload Component

'-----------------------
'Result data display area
'----------------------- 

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>
