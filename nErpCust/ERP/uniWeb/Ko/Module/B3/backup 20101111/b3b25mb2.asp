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
		
	ReDim IG1_Import_Group(UBound(arrRows,1),25)

	For LngRow = 0 To LngMaxRow - 1

		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))												'☜: Row 의 상태 

		IG1_Import_Group(LngRow, C_IG1_I2_select_char) = UCase(Trim(arrCols(0)))
		IG1_Import_Group(LngRow, C_IG1_I5_item_cd) = UCase(Trim(arrCols(2)))
		IG1_Import_Group(LngRow, C_IG1_I1_prc_ctrl_indctr) = UCase(Trim(arrCols(3)))
			
		If UCase(Trim(arrCols(5))) = "N" And UniConvNum(arrCols(4),0) = 0 And arrCols(3) = "S" Then
			Call DisplayMsgBox("970022", VBOKOnly, "단가", 0, I_MKSCRIPT)
			Call SheetFocus(arrCols(1),"parent.C_UnitPrice",I_MKSCRIPT)
			Response.End
		End If
		
		If arrCols(3) = "S" Then
			IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum(arrCols(4),0)
			IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum("0",0)
		Else
			IG1_Import_Group(LngRow, C_IG1_I1_std_prc)			= UniConvNum("0",0)
			IG1_Import_Group(LngRow, C_IG1_I1_moving_avg_prc)	= UniConvNum(arrCols(4),0)
		End If
				
		IG1_Import_Group(LngRow, C_IG1_I5_phantom_flg)			= UCase(Trim(arrCols(5)))
		IG1_Import_Group(LngRow, C_IG1_I5_valid_from_dt)		= UniConvDate(arrCols(6))
		IG1_Import_Group(LngRow, C_IG1_I5_valid_to_dt)			= UniConvDate(arrCols(7))

		If UniConvDateToYYYYMMDD(arrCols(6),gServerDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(7),gServerDateFormat,"") Then
			Call DisplayMsgBox("970023", VBOKOnly, "유효종료일", "유효시작일", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
			Response.End
		End If

		If UniConvDateToYYYYMMDD(arrCols(6),gServerDateFormat,"") < UniConvDateToYYYYMMDD(arrCols(8),gServerDateFormat,"") Then
			Call DisplayMsgBox("970023", VBOKOnly, "유효시작일", "품목시작일", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidFromDt", I_MKSCRIPT)
			Response.End
		End If

		If UniConvDateToYYYYMMDD(arrCols(7),gServerDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(9),gServerDateFormat,"") Then
			Call DisplayMsgBox("970025", VBOKOnly, "유효종료일", "품목종료일", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
			Response.End
		End If

		IG1_Import_Group(LngRow, C_IG1_I3_vchar_11)				= Trim(arrCols(10))
		
        'If LngRow >= 49 Or LngRow = LngMaxRow - 1 Then							'⊙: 5개를 Group으로, 나머지 일때 
		'	Exit For
		'End If
	
	Next
		
	I1_B_Plant_Plant_Cd	= Trim(Request("txtPlantCd"))
	I2_B_Item_Item_Cd	= Trim(Request("txtItemCd1"))
	I3_B_Plant_Plant_Cd = Trim(Request("hPlantCd"))
		
	Set oPP1G301 = Server.CreateObject("PP1G301.cPCopyItemMaster")    

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