<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211MB2
'*  4. Program Name         : 품목별 검사기준 등록 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG110
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd 
	
Dim PQBG110																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim LngMaxRow
Dim arrRowVal								'.☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt								'☜: Group Count
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strRoutNo
Dim strOprNo
dim IG1_import_group
Dim iErrorPosition
Dim LngRow

Const IG1_I1_row_num = 0
Const IG1_I1_insp_item_cd = 1
Const IG1_I1_select_char = 2
Const IG1_I1_measmt_equipmt_cd = 3
Const IG1_I1_insp_method_cd = 4
Const IG1_I1_weight_cd = 5
Const IG1_I1_insp_spec = 6
Const IG1_I1_usl = 7
Const IG1_I1_lsl = 8
Const IG1_I1_measmt_unit_cd = 9
Const IG1_I1_ucl = 10
Const IG1_I1_lcl = 11
Const IG1_I1_mthd_of_cl_cal = 12
Const IG1_I1_calculated_qty = 13
Const IG1_I1_insp_order = 14
Const IG1_I1_insp_unit_indctn = 15
Const IG1_I1_insp_process_desc = 16
Const IG1_I1_remark = 17
    
Const C_SHEETMAXROWS_D = 100
	
LngMaxRow = CInt(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 

If CInt(Request("txtFlgMode")) = OPMD_UMODE	Then	'☜: 저장시 Create/Update 판별 
	strPlantCd		= UCase(Request("hPlantCd"))
	strItemCd		= Request("hItemCd")
	strInspClassCd	= Request("hInspClassCd")
	strRoutNo		= Request("hRoutNo")
	strOprNo		= Request("hOprNo")
		
Else	
	strPlantCd		= UCase(Request("txtPlantCd"))
	strItemCd		= UCase(Request("txtItemCd"))
	strInspClassCd	= UCase(Request("cboInspClassCd"))
	strRoutNo		= UCase(Request("txtRoutNo"))
	strOprNo		= UCase(Request("txtOprNo"))
End If			
			
Set PQBG110 = Server.CreateObject("PQBG110.cQMtInspStdItemSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
		
lGrpCnt  = 0
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	Redim IG1_import_group(LngMaxRow, 17)
	For LngRow = 1 To LngMaxRow
		arrColVal = Split(arrRowVal(LngRow-1), gColSep)
		lGrpCnt = lGrpCnt +1														'☜: Group Count
		strStatus = arrColVal(0)				
		Select Case strStatus
			Case "C"	
				IG1_import_group(LngRow,IG1_I1_insp_item_cd) = UCase(arrColVal(1))
				IG1_import_group(LngRow,IG1_I1_insp_order) = arrColVal(2)	
				IG1_import_group(LngRow,IG1_I1_insp_method_cd) = UCase(arrColVal(3))
				IG1_import_group(LngRow,IG1_I1_insp_unit_indctn) = arrColVal(4)
				IG1_import_group(LngRow,IG1_I1_weight_cd) = arrColVal(5)
				IG1_import_group(LngRow,IG1_I1_insp_spec) = arrColVal(6)	
				if arrColVal(7) <> "" Then
					IG1_import_group(LngRow,IG1_I1_lsl) = UNIConvNum(arrColVal(7), 0)
				End If
				if arrColVal(8) <> "" Then
					IG1_import_group(LngRow,IG1_I1_usl) = UNIConvNum(arrColVal(8), 0)
				End If
				IG1_import_group(LngRow,IG1_I1_mthd_of_cl_cal) = arrColVal(9)	
				if arrColVal(10) <> "" Then
					IG1_import_group(LngRow,IG1_I1_calculated_qty) = UNIConvNum(arrColVal(10), 0)
				End If
				if arrColVal(11) <> "" Then
					IG1_import_group(LngRow,IG1_I1_lcl) = UNIConvNum(arrColVal(11), 0)
				End If
				if arrColVal(12) <> "" Then
					IG1_import_group(LngRow,IG1_I1_ucl) = UNIConvNum(arrColVal(12), 0)
				End If
				IG1_import_group(LngRow,IG1_I1_measmt_equipmt_cd) = UCase(arrColVal(13))
				IG1_import_group(LngRow,IG1_I1_measmt_unit_cd) = arrColVal(14)
				IG1_import_group(LngRow,IG1_I1_insp_process_desc) = arrColVal(15)
				IG1_import_group(LngRow,IG1_I1_remark) = arrColVal(16)
				IG1_import_group(LngRow,IG1_I1_row_num) = arrColVal(15)
				IG1_import_group(LngRow,IG1_I1_select_char) = "C"
			Case "U"	
				IG1_import_group(LngRow,IG1_I1_insp_item_cd) = UCase(arrColVal(1))
				IG1_import_group(LngRow,IG1_I1_insp_order) = arrColVal(2)	
				IG1_import_group(LngRow,IG1_I1_insp_method_cd) = UCase(arrColVal(3))
				IG1_import_group(LngRow,IG1_I1_insp_unit_indctn) = arrColVal(4)
				IG1_import_group(LngRow,IG1_I1_weight_cd) = arrColVal(5)
				IG1_import_group(LngRow,IG1_I1_insp_spec) = arrColVal(6)	
				if arrColVal(7) <> "" Then
					IG1_import_group(LngRow,IG1_I1_lsl) = UNIConvNum(arrColVal(7), 0)
				End If
				if arrColVal(8) <> "" Then
					IG1_import_group(LngRow,IG1_I1_usl) = UNIConvNum(arrColVal(8), 0)
				End If
				IG1_import_group(LngRow,IG1_I1_mthd_of_cl_cal) = arrColVal(9)	
				if arrColVal(10) <> "" Then
					IG1_import_group(LngRow,IG1_I1_calculated_qty) = UNIConvNum(arrColVal(10), 0)
				End If
				if arrColVal(11) <> "" Then
					IG1_import_group(LngRow,IG1_I1_lcl) = UNIConvNum(arrColVal(11), 0)
				End If
				if arrColVal(12) <> "" Then
					IG1_import_group(LngRow,IG1_I1_ucl) = UNIConvNum(arrColVal(12), 0)
				End If
				IG1_import_group(LngRow,IG1_I1_measmt_equipmt_cd) = UCase(arrColVal(13))
				IG1_import_group(LngRow,IG1_I1_measmt_unit_cd) = arrColVal(14)
				IG1_import_group(LngRow,IG1_I1_insp_process_desc) = arrColVal(15)
				IG1_import_group(LngRow,IG1_I1_remark) = arrColVal(16)
				IG1_import_group(LngRow,IG1_I1_row_num) = arrColVal(15)
				IG1_import_group(LngRow,IG1_I1_select_char) = "U"	
			Case "D"
				IG1_import_group(LngRow,IG1_I1_insp_item_cd) = UCase(arrColVal(1))
				IG1_import_group(LngRow,IG1_I1_row_num) = arrColVal(2)	
				IG1_import_group(LngRow,IG1_I1_select_char) = "D"
	 	End Select
				 	
	Next
					
	Call PQBG110.Q_MT_INSP_STD_BY_ITEM_SVR (gStrGlobalCollection, strPlantCd, strItemCd, _
											strInspClassCd, strRoutNo, strOprNo, _
											IG1_import_group, iErrorPosition)


	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		If iErrorPosition <> "" Then	
			Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
			Set PQBG110 = Nothing
			Response.End
		End If
		Response.End
	End If
End If
Set PQBG110 = Nothing                                                   '☜: Unload Comproxy
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	Dim strHTML
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
</Script>