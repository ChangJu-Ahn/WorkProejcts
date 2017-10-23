<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1412mb2.asp
'*  4. Program Name         : 일괄변경 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf


Dim oPP1G421																	'☆ : 저장용 ComProxy Dll 사용 변수 

Dim StrNextKey											'⊙: 다음 값 
Dim lgStrPrevKey										'⊙: 이전 값 
Dim strMode												'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngMaxRow											'⊙: 현재 그리드의 최대Row
Dim intFlgMode
Dim LngRow
Dim LngIdx

Dim I1_Plant_Cd
Dim I2_Child_Item_Cd
Dim I3_Bom_Type
Dim IG1_Bom_Detail

Dim arrCols, arrRows									'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus											'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt												'☜: Group Count
Dim strCode												'⊙: Lookup 용 리턴 변수 

ReDim IG1_Bom_Detail(12)
Const C_Prnt_Item_Cd = 0
Const C_Child_Item_Seq = 1
Const C_Child_Item_Qty = 2
Const C_Child_Unit = 3
Const C_Prnt_Item_Qty = 4
Const C_Prnt_Unit = 5
Const C_Safety_Lt = 6
Const C_Loss_Rate = 7
Const C_Supply_Type = 8
Const C_Valid_From_Dt = 9
Const C_Valid_To_Dt = 10
Const C_Ecn_No = 11
Const C_Reason_Cd = 12
Const C_Ecn_Desc = 13
Const C_Remark = 14

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 
	intFlgMode = CInt(Request("txtFlgMode"))	
	LngMaxRow = CInt(Request("txtMaxRows"))									'☜: 최대 업데이트된 갯수 
    lgStrPrevKey = Trim(Request("lgStrPrevKey"))

	If Trim(Request("hPlantCd")) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	If Trim(Request("hItemCd")) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	If Trim(Request("hBomType")) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	I1_Plant_Cd = UCase(Trim(Request("hPlantCd")))
	I2_Child_Item_Cd = UCase(Trim(Request("hItemCd")))
	I3_Bom_Type = UCase(Trim(Request("hBomType")))
	
    arrRows = Split(Request("txtSpread"), gRowSep)							'☆: Spread Sheet 내용을 담고 있는 Element명 
    ReDim IG1_Bom_Detail(UBound(arrRows,1), 14)

	For LngRow = 0 To LngMaxRow - 1
		arrCols = Split(arrRows(LngRow), gColSep)
		
		If CDbl(arrCols(4)) = 0 Then
			Call DisplayMsgBox("970022", VBOKOnly, "자품목기준수", "0", I_MKSCRIPT)	
			Call SheetFocus(arrCols(1), "parent.C_ChildItemQty", I_MKSCRIPT)
			Response.End
		End If
		
		If CDbl(arrCols(6)) = 0 Then
			Call DisplayMsgBox("970022", VBOKOnly, "모품목기준수", "0", I_MKSCRIPT)	
			Call SheetFocus(arrCols(1), "parent.C_PrntItemQty", I_MKSCRIPT)
			Response.End
		End If
		
		If UniConvDateToYYYYMMDD(arrCols(11),gDateFormat,"") > UniConvDateToYYYYMMDD(arrCols(12),gDateFormat,"") Then
			Call DisplayMsgBox("970023", VBOKOnly, "종료일", "시작일", I_MKSCRIPT)
			Call SheetFocus(arrCols(1), "parent.C_ValidFromDt", I_MKSCRIPT)
			Response.End
		End If

		IG1_Bom_Detail(LngRow, C_Prnt_Item_Cd)	= UCase(Trim(arrCols(2)))
		IG1_Bom_Detail(LngRow, C_Child_Item_Seq)= UCase(Trim(arrCols(3)))
		IG1_Bom_Detail(LngRow, C_Child_Item_Qty)= UniConvNum(arrCols(4), 0)
		IG1_Bom_Detail(LngRow, C_Child_Unit)	= UCase(Trim(arrCols(5)))
		IG1_Bom_Detail(LngRow, C_Prnt_Item_Qty)	= UniConvNum(arrCols(6), 0)
		IG1_Bom_Detail(LngRow, C_Prnt_Unit)		= UCase(Trim(arrCols(7)))
		IG1_Bom_Detail(LngRow, C_Safety_Lt)		= UniConvNum(arrCols(8), 0)
		IG1_Bom_Detail(LngRow, C_Loss_Rate)		= UniConvNum(arrCols(9), 0)
		IG1_Bom_Detail(LngRow, C_Supply_Type)	= UCase(Trim(arrCols(10)))
		IG1_Bom_Detail(LngRow, C_Valid_From_Dt)	= UniConvDateToYYYYMMDD(arrCols(11),gDateFormat,"-")
		IG1_Bom_Detail(LngRow, C_Valid_To_Dt)	= UniConvDateToYYYYMMDD(arrCols(12),gDateFormat,"-")
		IG1_Bom_Detail(LngRow, C_Ecn_No)		= UCase(Trim(arrCols(13)))
		IG1_Bom_Detail(LngRow, C_Reason_Cd)		= UCase(Trim(arrCols(14)))
		IG1_Bom_Detail(LngRow, C_Ecn_Desc)		= UCase(Trim(arrCols(15)))
		IG1_Bom_Detail(LngRow, C_Remark)		= arrCols(16)
		
	Next

	Set oPP1G421 = Server.CreateObject("PP1G421.cPBomMassChange")

	If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

	Call oPP1G421.P_BOM_MASS_CHANGE(gStrGlobalCollection, _
									I1_Plant_Cd, _
									I2_Child_Item_Cd, _
									I3_Bom_Type, _
									IG1_Bom_Detail)

	If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G421 = Nothing		                                        '☜: Unload Comproxy DLL
       Response.End
    End If

	Set oPP1G421 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr					
	
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

%>
