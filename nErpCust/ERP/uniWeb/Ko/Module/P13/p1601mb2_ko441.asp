<%@LANGUAGE = VBScript%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Dim oPP1G301

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim intFlgMode
Dim StrNextKey																'⊙: 다음 값 
Dim lgStrPrevKey															'⊙: 이전 값 
Dim LngMaxRow																'⊙: 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount 
Dim strPlantCd
Dim strItemCd

Dim I1_B_Item_Item_Cd
Dim I2_B_Plant_Plant_Cd
Dim IG1_Import_Group

Const C_IG1_I1_prc_ctrl_indctr = 0
Const C_IG1_I1_moving_avg_prc = 1
Const C_IG1_I1_std_prc = 2
Const C_IG1_I2_select_char = 3
Const C_IG1_I3_vchar_11 = 4
Const C_IG1_I4_item_group_cd = 5
Const C_IG1_I5_item_cd = 6
Const C_IG1_I5_item_nm = 7
Const C_IG1_I5_formal_nm = 8
Const C_IG1_I5_spec = 9
Const C_IG1_I5_item_acct = 10
Const C_IG1_I5_item_class = 11
Const C_IG1_I5_hs_cd = 12
Const C_IG1_I5_hs_unit = 13
Const C_IG1_I5_unit_weight = 14
Const C_IG1_I5_unit_of_weight = 15
Const C_IG1_I5_basic_unit = 16
Const C_IG1_I5_phantom_flg = 17
Const C_IG1_I5_draw_no = 18
Const C_IG1_I5_blanket_pur_flg = 19
Const C_IG1_I5_base_item_cd = 20
Const C_IG1_I5_valid_flg = 21
Const C_IG1_I5_valid_from_dt = 22
Const C_IG1_I5_valid_to_dt = 23
Const C_IG1_I5_vat_type = 24
Const C_IG1_I5_vat_rate = 25
Const C_IG1_I5_gross_weight = 26
Const C_IG1_I5_gross_unit = 27
Const C_IG1_I5_cbm = 28
Const C_IG1_I5_cbm_description = 29

Call HideStatusWnd

On Error Resume Next	
Err.Clear

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 
	intFlgMode = CInt(Request("txtFlgMode"))
	LngMaxRow = CInt(Request("txtMaxRows"))	
    
    Dim arrCols, arrRows														'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus															'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																'☜: Group Count
	Dim strCode																'⊙: Lookup 용 리턴 변수 
	
		
	arrRows = Split(Request("txtSpread"), gRowSep)							'☆: Spread Sheet 내용을 담고 있는 Element명 

	ReDim IG1_Import_Group(UBound(arrRows,1),29)

	For LngRow = 0 To LngMaxRow - 1

		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))												'☜: Row 의 상태 

		Select Case strStatus

		    Case "C"		
			
				IG1_Import_Group(LngRow, C_IG1_I2_select_char) = UCase(Trim(arrCols(0)))
				IG1_Import_Group(LngRow, C_IG1_I5_item_cd) = UCase(Trim(arrCols(2)))
				IG1_Import_Group(LngRow, C_IG1_I1_prc_ctrl_indctr) = UCase(Trim(arrCols(3)))
				
				'If UCase(Trim(arrCols(5))) = "N" And UniConvNum(arrCols(4),0) = 0 And arrCols(3) = "S" Then
				'	Call DisplayMsgBox("970022", VBOKOnly, "단가", 0, I_MKSCRIPT)
				'	Call SheetFocus(arrCols(1),"parent.C_UnitPrice",I_MKSCRIPT)
				'	Response.End
				'End If
				If arrCols(3) = "S" Then
					IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum(arrCols(4),0)
					IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum("0",0)
				Else
					IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum("0",0)
					IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum("0",0)
					'IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum(arrCols(4),0)
				End If
				
				IG1_Import_Group(LngRow, C_IG1_I5_phantom_flg)			= UCase(Trim(arrCols(5)))
				IG1_Import_Group(LngRow, C_IG1_I5_valid_from_dt)		= UniConvDate(arrCols(6))
				IG1_Import_Group(LngRow, C_IG1_I5_valid_to_dt)			= UniConvDate(arrCols(7))
				
				If UniConvDateToYYYYMMDD(arrCols(6),gDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(7),gDateFormat,"") Then
					Call DisplayMsgBox("970023", VBOKOnly, "유효종료일", "유효시작일", I_MKSCRIPT)
					Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
					Response.End
				End If
				
				If UniConvDateToYYYYMMDD(arrCols(6),gDateFormat,"") < UniConvDateToYYYYMMDD(arrCols(8),gDateFormat,"") Then
					Call DisplayMsgBox("970023", VBOKOnly, "유효시작일", "품목시작일", I_MKSCRIPT)
					Call SheetFocus(arrCols(1), "parent.C_IBPValidFromDt", I_MKSCRIPT)
					Response.End
				End If

				If UniConvDateToYYYYMMDD(arrCols(7),gDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(9),gDateFormat,"") Then
					Call DisplayMsgBox("970025", VBOKOnly, "유효종료일", "품목종료일", I_MKSCRIPT)
					Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
					Response.End
				End If
				
				IG1_Import_Group(LngRow, C_IG1_I3_vchar_11)				= Trim(arrCols(10))
			
			Case "I"

				IG1_Import_Group(LngRow, C_IG1_I2_select_char) = UCase(Trim(arrCols(0)))
				IG1_Import_Group(LngRow, C_IG1_I5_item_cd) = UCase(Trim(arrCols(2)))
				IG1_Import_Group(LngRow, C_IG1_I5_item_nm) = arrCols(3)
				IG1_Import_Group(LngRow, C_IG1_I5_formal_nm) = arrCols(4)
				IG1_Import_Group(LngRow, C_IG1_I5_item_acct) = Trim(arrCols(5))
				IG1_Import_Group(LngRow, C_IG1_I5_basic_unit) = UCase(Trim(arrCols(6)))
				IG1_Import_Group(LngRow, C_IG1_I4_item_group_cd) = UCase(Trim(arrCols(7)))
				IG1_Import_Group(LngRow, C_IG1_I5_phantom_flg) = UCase(Trim(arrCols(8)))
				IG1_Import_Group(LngRow, C_IG1_I5_blanket_pur_flg) = UCase(Trim(arrCols(9)))
				IG1_Import_Group(LngRow, C_IG1_I5_base_item_cd) = UCase(Trim(arrCols(10)))
				IG1_Import_Group(LngRow, C_IG1_I5_item_class) = UCase(Trim(arrCols(11)))
				IG1_Import_Group(LngRow, C_IG1_I5_valid_flg) = UCase(Trim(arrCols(12)))
				IG1_Import_Group(LngRow, C_IG1_I5_spec) = arrCols(13)
				IG1_Import_Group(LngRow, C_IG1_I5_unit_weight) =UniConvNum(arrCols(14),0)
				IG1_Import_Group(LngRow, C_IG1_I5_unit_of_weight) = UCase(Trim(arrCols(15)))
				IG1_Import_Group(LngRow, C_IG1_I5_draw_no) = arrCols(16)
				IG1_Import_Group(LngRow, C_IG1_I5_hs_cd) = arrCols(17)
				IG1_Import_Group(LngRow, C_IG1_I5_hs_unit) = Trim(arrCols(18))
				IG1_Import_Group(LngRow, C_IG1_I5_valid_from_dt) = UniConvDate(arrCols(19))
				IG1_Import_Group(LngRow, C_IG1_I5_valid_to_dt) = UniConvDate(arrCols(20))		
				
				If UniConvDateToYYYYMMDD(arrCols(19),gDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(20),gDateFormat,"") Then
					Call DisplayMsgBox("970023", VBOKOnly, "유효종료일", "유효시작일", I_MKSCRIPT)
					Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
					Response.End
				End If
				
				IG1_Import_Group(LngRow, C_IG1_I1_prc_ctrl_indctr) = Trim(arrCols(21))
				
				'If UCase(Trim(arrCols(8))) = "N" And UniConvNum(arrCols(22),0) = 0 Then
				'	Call DisplayMsgBox("970022", VBOKOnly, "단가", 0, I_MKSCRIPT)
				'	Call SheetFocus(arrCols(1),"parent.C_UnitPrice",I_MKSCRIPT)
				'	Response.End
				'End If
				
				If arrCols(21) = "S" Then
					IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum(arrCols(22),0)
					IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum("0",0)
				Else
					IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum("0",0)
					IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum(arrCols(22),0)
				End If
				
				IG1_Import_Group(LngRow, C_IG1_I5_vat_type)				= Trim(arrCols(23))
				IG1_Import_Group(LngRow, C_IG1_I5_vat_rate)				= UniConvNum(arrCols(24),0)
				IG1_Import_Group(LngRow, C_IG1_I3_vchar_11)				= Trim(arrCols(25))
				
				'appended in 2003-07-01
				IG1_Import_Group(LngRow, C_IG1_I5_gross_weight) =UniConvNum(arrCols(26),0)
				IG1_Import_Group(LngRow, C_IG1_I5_gross_unit) = UCase(Trim(arrCols(27)))
				IG1_Import_Group(LngRow, C_IG1_I5_cbm) =UniConvNum(arrCols(28),0)
				IG1_Import_Group(LngRow, C_IG1_I5_cbm_description) = Trim(arrCols(29))
				
		End Select        

	Next
	
	I1_B_Plant_Plant_Cd	= Trim(Request("txtPlantCd"))
	I2_B_Item_Item_Cd	= Trim(Request("txtItemCd1"))
	I3_B_Plant_Plant_Cd = Trim(Request("txtPlantCd1"))

	Set oPP1G301 = Server.CreateObject("PP1G301_KO441.cPCopyItemMaster")            '20080507::HANC

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End 
	End If	
	
	Call oPP1G301.P_COPY_ITEM_MASTER(gStrGlobalCollection, _
									IG1_Import_Group, _
									I1_B_Plant_Plant_Cd, _
									I2_B_Item_Item_Cd, _
									I3_B_Plant_Plant_Cd)
									
	If CheckSYSTEMError(Err,True) = True Then
	   Set oPP1G301 = Nothing
	   Response.End		
	End If

	Set oPP1G301 = Nothing

	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "	With parent " & vbCr																		'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "		.DbSaveOk " & vbCr
	Response.Write "	End With " & vbCr
	Response.Write "</Script> " & vbCr


'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>