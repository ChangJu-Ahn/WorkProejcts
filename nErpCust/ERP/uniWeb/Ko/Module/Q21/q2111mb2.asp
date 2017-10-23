<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2111MB2
'*  4. Program Name         : 검사등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
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
							
On Error Resume Next
Call HideStatusWnd 
	
Dim PQIG010																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt	
Dim strUserId
Dim iCommand
Dim I3_q_inspection_result
Dim iErrorPosition
	
Dim IG1_import_group
	
Const IG1_IefSuppliedCommand = 0
Const IG1_InspItemCd = 1
Const IG1_InspSeries = 2
Const IG1_SampleQty = 3
Const IG1_AccptDecisionQty = 4
Const IG1_RejtDecisionQty = 5
Const IG1_AccptDecisionDiscreate = 6
Const IG1_MaxDefectRatio = 7
Const IG1_MeasmtEquipmtCd = 8
Const IG1_MeasmtUnitCd = 9
Const IG1_Row = 10
	
lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
LngMaxRow = CInt(Request("txtMaxRows"))	
strUserId = Request("txtInsrtUserId")
		
Set PQIG010 = Server.CreateObject("PQIG010.cQMtInspResultSvr")

If CheckSystemError(Err,True) Then
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
    
Redim I3_q_inspection_result(3)
    
I3_q_inspection_result(0) = 1
I3_q_inspection_result(1) = UNIConvNum(Request("txtLotSize"), 0)
I3_q_inspection_result(2) = "R"
I3_q_inspection_result(3) = UCase(Trim(Request("txtPlantCd")))
	
If lgIntFlgMode = OPMD_CMODE Then
	iCommand   = "C"					
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iCommand   = "U"					
End If
	
'Detail Import
		
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	LngMaxRow = Ubound(arrRowVal)
	Redim IG1_Import_Group(LngMaxRow ,10)
				
	For LngRow = 0 To LngMaxRow - 1
		arrColVal = Split(arrRowVal(LngRow), gColSep)
																'☜: Group Count
		strStatus = UCase(arrColVal(0)	)
			
		IG1_import_group(LngRow,IG1_InspItemCd) = arrColVal(1)
		IG1_import_group(LngRow,IG1_InspSeries) = UniConvNum(arrColVal(2),0)
		IG1_import_group(LngRow,IG1_IefSuppliedCommand) = strStatus	
		
		If strStatus = "C" or strStatus = "U" Then				'☜: Row 의 상태 
			IG1_import_group(LngRow,IG1_SampleQty) = UniConvNum(arrColVal(3),0)	
			IG1_import_group(LngRow,IG1_AccptDecisionQty) = UniConvNum(arrColVal(4), 0)	
			IG1_import_group(LngRow,IG1_RejtDecisionQty) = UniConvNum(arrColVal(5), 0)

			If arrColVal(6) <> "" Then 
				IG1_import_group(LngRow,IG1_AccptDecisionDiscreate) = UNIConvNum(arrColVal(6), 0)
			Else
				IG1_import_group(LngRow,IG1_AccptDecisionDiscreate) = ""
			End If

			If arrColVal(7) <> "" Then 
				IG1_import_group(LngRow,IG1_MaxDefectRatio) = UNIConvNum(arrColVal(7), 0)
			Else
				IG1_import_group(LngRow,IG1_MaxDefectRatio) = ""
			End If			
			
			IG1_import_group(LngRow,IG1_MeasmtEquipmtCd) = arrColVal(8)	
			IG1_import_group(LngRow,IG1_MeasmtUnitCd) = arrColVal(9)
			IG1_import_group(LngRow,IG1_Row) = arrColVal(10)
		Else
			IG1_import_group(LngRow,IG1_Row) = arrColVal(3)
	 	End If
	Next
		
	Dim strtxtInspReqNo2
	strtxtInspReqNo2 = Request("txtInspReqNo2")
		
	Call PQIG010.Q_MAINT_INSP_RESULT_SVR(gStrGlobalCollection,iCommand,UCase(Trim(strtxtInspReqNo2)),I3_q_inspection_result,IG1_import_group,iErrorPosition)
		
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
		Set PQIG010 = Nothing 
		Response.End
	End if

			
End If
	
Set PQIG010 = Nothing                                                   '☜: Unload Comproxy
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