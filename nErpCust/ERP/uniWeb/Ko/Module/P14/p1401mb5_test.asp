<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1401mb5.asp	
'*  4. Program Name         : BOM Detail Save
'*  5. Program Desc         :
'*  6. Component List       : PP1S405.cPMngBomDtl
'*  7. Modified date(First) : 2000/05/2
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd

On Error Resume Next
Err.Clear																		'☜: Protect system from crashing

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim pPP1S405																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iCommandSent, I1_select_char, I2_plant_cd, I3_prnt_item_cd, I4_prnt_bom_no ,I5_child_item_cd
Dim I6_p_bom_header , I7_p_bom_detail

Const P198_I6_bom_no = 0
Const P198_I6_description = 1
Const P198_I6_valid_from_dt = 2
Const P198_I6_valid_to_dt = 3
Const P198_I6_drawing_path = 4

Const P198_I7_child_item_seq = 0
Const P198_I7_child_bom_no = 1
Const P198_I7_prnt_item_qty = 2
Const P198_I7_prnt_item_unit = 3
Const P198_I7_child_item_qty = 4
Const P198_I7_child_item_unit = 5
Const P198_I7_loss_rate = 6
Const P198_I7_safety_lt = 7
Const P198_I7_supply_type = 8
Const P198_I7_bom_flg = 9
Const P198_I7_valid_from_dt = 10
Const P198_I7_valid_to_dt = 11
Const P198_I7_ecn_no = 12
Const P198_I7_ecn_desc = 13
Const P198_I7_reason_cd = 14
Const P198_I7_remark = 15
            
If Request("txtPlantCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

If Request("txtPrntItemCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

If Request("txtDtlMode") <> "D" Then
	If Request("txtItemCd1") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
End If
	
'-----------------------
'Data manipulate area
'-----------------------
    
Redim I6_p_bom_header(P198_I6_drawing_path)
Redim I7_p_bom_detail(P198_I7_remark)
	
I2_plant_cd = UCase(Request("txtPlantCd"))
	
'------------------------------------------------------
'반제품인 경우 자신의 BOM Header가 처음 등록될 때 
'------------------------------------------------------
If Request("txtDtlMode") <> "D" Then
	If Request("txtHdrMode") <> "" Then
		If Request("txtHdrMode") = "C" Then
			iCommandSent = "CREATE"
		ElseIf Request("txtHdrMode") = "U" Then
			iCommandSent = "UPDATE"
		End If

		I6_p_bom_header(P198_I6_bom_no)			= UCase(Request("txtBomNo1"))
		I7_p_bom_detail(P198_I7_child_bom_no)	= UCase(Request("txtBomNo1"))
		I6_p_bom_header(P198_I6_description)	= Request("txtBOMDesc")
		I6_p_bom_header(P198_I6_drawing_path)	= Request("txtDrawPath")
			
		If Len(Trim(Request("txtValidFromDt"))) Then
			If UniConvDate(Request("txtValidFromDt")) = "" Then	 
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
				Response.End	
			Else
				I6_p_bom_header(P198_I6_valid_from_dt) = UniConvDate(Request("txtValidFromDt"))
			End If
		End If
	
		If Len(Trim(Request("txtValidToDt"))) Then
			If UniConvDate(Request("txtValidToDt")) = "" Then	 
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
				Response.End	
			Else
				I6_p_bom_header(P198_I6_valid_to_dt) = UniConvDate(Request("txtValidToDt"))
			End If
		End If
		
	End If

	'-------------------------------------------------------
	' 자품목 정보를 생성 
	'-------------------------------------------------------
	If Request("txtDtlMode") <> "" Then
		If Request("txtDtlMode") = "C" Then
			I1_select_char	= "C"
		ElseIf Request("txtDtlMode") = "U" Then
			I1_select_char	= "U"
		End If
			
		'-------------------------------------------------------
		'Detail생성을 위한 모품목의 BOM Header정보 
		'-------------------------------------------------------
'20071207::hanc	I3_prnt_item_cd		= UCase(Request("txtPrntItemCd"))
        I3_prnt_item_cd = Trim(Mid(UCase(Request("txtPrntItemCd")), 1, InStr(1, UCase(Request("txtPrntItemCd")) & "(", "(")-1))
		I4_prnt_bom_no		= UCase(Request("txtPrntBomNo"))
		'-------------------------------------------------------
		'BOM Detail 정보를 주어야 할 때 
		'-------------------------------------------------------
			
		I5_child_item_cd		= UCase(Request("txtItemCd1"))
		
		I7_p_bom_detail(P198_I7_child_bom_no)	= UCase(Request("txtPrntBomNo"))
		I7_p_bom_detail(P198_I7_child_item_seq)	= CInt(Request("txtItemSeq"))
		I7_p_bom_detail(P198_I7_child_item_qty)	= UniConvNum(Request("txtChildItemQty"),0)
		I7_p_bom_detail(P198_I7_child_item_unit )	= UCase(Trim(Request("txtChildItemUnit"))) 
		I7_p_bom_detail(P198_I7_prnt_item_qty)	= UniConvNum(Request("txtPrntItemQty"),0)
		I7_p_bom_detail(P198_I7_prnt_item_unit)	= UCase(Trim(Request("txtPrntItemUnit")))
		I7_p_bom_detail(P198_I7_safety_lt)	= UniConvNum(Request("txtSafetyLt"),0)
		I7_p_bom_detail(P198_I7_loss_rate)	= UniConvNum(Request("txtLossRate"),0)
	
		If Request("rdoSupplyFlg") <> "" Then			
			I7_p_bom_detail(P198_I7_supply_type)	= Request("rdoSupplyFlg")
		Else
			I7_p_bom_detail(P198_I7_supply_type)	= "F"
		End If
			
		I7_p_bom_detail(P198_I7_bom_flg)  				= UniConvNum(Request("cboBomFlg"),0)
			
		'---------------------------------------------
		' 추가일 : 2001-03-13
		'---------------------------------------------							
		I7_p_bom_detail(P198_I7_valid_from_dt) = UniConvDate(Request("txtValidFromDt1"))
		I7_p_bom_detail(P198_I7_valid_to_dt) = UniConvDate(Request("txtValidToDt1"))
		I7_p_bom_detail(P198_I7_ecn_no) = UCase(Request("txtECNNo1"))
		I7_p_bom_detail(P198_I7_ecn_desc) = Request("txtECNDesc1")
		I7_p_bom_detail(P198_I7_reason_cd) = Request("txtReasonCd1")
		I7_p_bom_detail(P198_I7_remark)	= Request("txtRemark")
		
	End If
Else
		
	I1_select_char	= "D"
	'-------------------------------------------------------
	'Detail삭제를 위한 모품목의 BOM Header정보 
	'-------------------------------------------------------
'20071210::hanc	I3_prnt_item_cd		= UCase(Request("txtPrntItemCd"))
    I3_prnt_item_cd = Trim(Mid(UCase(Request("txtPrntItemCd")), 1, InStr(1, UCase(Request("txtPrntItemCd")) & "(", "(")-1))
	I4_prnt_bom_no		= UCase(Request("txtPrntBomNo"))
	'-------------------------------------------------------
	'BOM Detail 정보를 주어야 할 때 (History 때문에 아래 변경(추가) kjpark(20030428)
	'-------------------------------------------------------
			 	
'	I7_p_bom_detail(P198_I7_child_item_seq)	= CInt(Request("txtItemSeq"))
	
				
		I5_child_item_cd		= UCase(Request("txtItemCd1"))
		
		I7_p_bom_detail(P198_I7_child_bom_no)	= UCase(Request("txtPrntBomNo"))
		I7_p_bom_detail(P198_I7_child_item_seq)	= CInt(Request("txtItemSeq"))
		I7_p_bom_detail(P198_I7_child_item_qty)	= UniConvNum(Request("txtChildItemQty"),0)
		I7_p_bom_detail(P198_I7_child_item_unit )	= UCase(Trim(Request("txtChildItemUnit"))) 
		I7_p_bom_detail(P198_I7_prnt_item_qty)	= UniConvNum(Request("txtPrntItemQty"),0)
		I7_p_bom_detail(P198_I7_prnt_item_unit)	= UCase(Trim(Request("txtPrntItemUnit")))
		I7_p_bom_detail(P198_I7_safety_lt)	= UniConvNum(Request("txtSafetyLt"),0)
		I7_p_bom_detail(P198_I7_loss_rate)	= UniConvNum(Request("txtLossRate"),0)
	
		If Request("rdoSupplyFlg") <> "" Then			
			I7_p_bom_detail(P198_I7_supply_type)	= Request("rdoSupplyFlg")
		Else
			I7_p_bom_detail(P198_I7_supply_type)	= "F"
		End If
			
		I7_p_bom_detail(P198_I7_bom_flg)  				= UniConvNum(Request("cboBomFlg"),0)
			
		'---------------------------------------------
		' 추가일 : 2001-03-13
		'---------------------------------------------							
		I7_p_bom_detail(P198_I7_valid_from_dt) = UniConvDate(Request("txtValidFromDt1"))
		I7_p_bom_detail(P198_I7_valid_to_dt) = UniConvDate(Request("txtValidToDt1"))
		I7_p_bom_detail(P198_I7_ecn_no) = UCase(Request("txtECNNo1"))
		I7_p_bom_detail(P198_I7_ecn_desc) = Request("txtECNDesc1")
		I7_p_bom_detail(P198_I7_reason_cd) = Request("txtReasonCd1")
		I7_p_bom_detail(P198_I7_remark)	= Request("txtRemark")

End If 		
	    
'-----------------------
'Com Action Area
'-----------------------
Set pPP1S405 = Server.CreateObject("PP1S405.cPMngBomDtl")
	    
If CheckSYSTEMError(Err,True) = True Then
	Set pPP1S405 = Nothing		
	Response.End
End If

Call pPP1S405.P_MANAGE_BOM_DETAIL(gStrGlobalCollection, iCommandSent, I1_select_char, I2_plant_cd, _
		I3_prnt_item_cd, I4_prnt_bom_no ,I5_child_item_cd, I6_p_bom_header , I7_p_bom_detail)
	
If CheckSYSTEMError(Err, True) = True Then
	Set pPP1S405 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1S405 = Nothing      

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
