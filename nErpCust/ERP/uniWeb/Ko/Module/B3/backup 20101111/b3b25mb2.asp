<%@LANGUAGE = VBScript%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Dim oPP1G301

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim intFlgMode
Dim StrNextKey																'��: ���� �� 
Dim lgStrPrevKey															'��: ���� �� 
Dim LngMaxRow																'��: ���� �׸����� �ִ�Row
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

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 
	intFlgMode = CInt(Request("txtFlgMode"))
	LngMaxRow = CInt(Request("txtMaxRows"))	
    
    Dim arrCols, arrRows														'��: Spread Sheet �� ���� ���� Array ���� 
	Dim strStatus															'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
	Dim	lGrpCnt																'��: Group Count
	Dim strCode																'��: Lookup �� ���� ���� 
		
	arrRows = Split(Request("txtSpread"), gRowSep)							'��: Spread Sheet ������ ��� �ִ� Element�� 
		
	ReDim IG1_Import_Group(UBound(arrRows,1),25)

	For LngRow = 0 To LngMaxRow - 1

		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))												'��: Row �� ���� 

		IG1_Import_Group(LngRow, C_IG1_I2_select_char) = UCase(Trim(arrCols(0)))
		IG1_Import_Group(LngRow, C_IG1_I5_item_cd) = UCase(Trim(arrCols(2)))
		IG1_Import_Group(LngRow, C_IG1_I1_prc_ctrl_indctr) = UCase(Trim(arrCols(3)))
			
		If UCase(Trim(arrCols(5))) = "N" And UniConvNum(arrCols(4),0) = 0 And arrCols(3) = "S" Then
			Call DisplayMsgBox("970022", VBOKOnly, "�ܰ�", 0, I_MKSCRIPT)
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
			Call DisplayMsgBox("970023", VBOKOnly, "��ȿ������", "��ȿ������", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
			Response.End
		End If

		If UniConvDateToYYYYMMDD(arrCols(6),gServerDateFormat,"") < UniConvDateToYYYYMMDD(arrCols(8),gServerDateFormat,"") Then
			Call DisplayMsgBox("970023", VBOKOnly, "��ȿ������", "ǰ�������", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidFromDt", I_MKSCRIPT)
			Response.End
		End If

		If UniConvDateToYYYYMMDD(arrCols(7),gServerDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(9),gServerDateFormat,"") Then
			Call DisplayMsgBox("970025", VBOKOnly, "��ȿ������", "ǰ��������", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_IBPValidToDt", I_MKSCRIPT)
			Response.End
		End If

		IG1_Import_Group(LngRow, C_IG1_I3_vchar_11)				= Trim(arrCols(10))
		
        'If LngRow >= 49 Or LngRow = LngMaxRow - 1 Then							'��: 5���� Group����, ������ �϶� 
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
	Response.Write "	With parent " & vbCr																		'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "		.DbSaveOk " & vbCr
	Response.Write "	End With " & vbCr
	Response.Write "</Script> " & vbCr


'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
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