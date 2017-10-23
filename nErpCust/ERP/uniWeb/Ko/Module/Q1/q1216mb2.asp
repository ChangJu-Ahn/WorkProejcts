<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1216MB2
'*  4. Program Name         : 기준정보 복사 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG210
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
Call LoadBasisGlobalInf
	
On Error Resume Next
Call HideStatusWnd 
	
Dim PQBG210																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim LngMaxRow
Dim LngMaxRow2
Dim LngRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt								'☜: Group Count
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strRoutNo
Dim strOprNo
	
Dim IG1_import_group1
Dim IG1_import_group2
Dim iErrorPosition
	
	LngMaxRow = CInt(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 
	LngMaxRow2 = CInt(Request("txtMaxRows2"))					'☜: 최대 업데이트된 갯수 
	strPlantCd = UCase(Request("hPlantCd"))
	strItemCd = UCase(Request("hItemCd"))
	strInspClassCd = UCase(Request("hInspClassCd"))
	strRoutNo = Request("txtRoutNo")
	strOprNo = Request("txtOprNo")
		
	Set PQBG210 = Server.CreateObject("PQBG210.cQMtInspStdCopySvr")

	If CheckSYSTEMError(Err,True) Then
		Response.End
	End if
		
	lGrpCnt  = 0
	If Request("txtSpread") <> "" Then
		arrRowVal = Split(Request("txtSpread"), gRowSep)
		LngMaxRow = UBound(arrRowVal) - 1
		Redim IG1_Import_Group1(LngMaxRow)
			
		For LngRow = 0 To LngMaxRow
			arrColVal = Split(arrRowVal(LngRow), gColSep)
			IG1_import_group1(LngRow) = arrColVal(0)
		Next
	End If
		
	lGrpCnt  = 0
		
	If Request("txtSpread2") <> "" Then
		arrRowVal = Split(Request("txtSpread2"), gRowSep)
		LngMaxRow2 = UBound(arrRowVal) - 1
			
		Redim IG1_import_group2(LngMaxRow2,3)
		
		For LngRow = 0 To LngMaxRow2
			arrColVal = Split(arrRowVal(LngRow), gColSep)
			IG1_import_group2(LngRow,0) = LngRow + 1
			IG1_import_group2(LngRow,1) = arrColVal(0)
			If strInspClassCd = "P" Then
				IG1_import_group2(LngRow,2) = arrColVal(1)
				IG1_import_group2(LngRow,3) = arrColVal(2)
			Else
				IG1_import_group2(LngRow,2) = ""
				IG1_import_group2(LngRow,3) = ""
			End if
		Next 
		
		Call PQBG210.Q_MAINT_INSP_STAND_COPY_SVR(gStrGlobalCollection, _
												 strPlantCd, _
												 strItemCd, _
												 strInspClassCd, _
												 strRoutNo, _
												 strOprNo, _
												 IG1_import_group1, _
												 IG1_Import_Group2, _
												 iErrorPosition)
			
		If CheckSYSTEMError(Err,True) Then
			Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
			Set PQBG210 = Nothing
			Response.End
		End if
	End if
	Set PQBG210 = Nothing                                                   '☜: Unload Comproxy
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
		strHTML = "parent.frm1.vspdData2.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.SelLength = len(parent.frm1.vspdData2.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData2.SelLength = len(parent.frm1.vspdData2.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>