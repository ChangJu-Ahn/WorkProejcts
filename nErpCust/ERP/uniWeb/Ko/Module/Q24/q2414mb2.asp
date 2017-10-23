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
'*  3. Program ID           : Q2414MB2
'*  4. Program Name         : 불량원인등록 
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

Dim strinsp_class_cd
strinsp_class_cd = "S"	'@@@주의 
	
Dim PQIG170
Dim LngMaxRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt									'☜: Group Count
Dim strUserId
Dim strInspReqNo
Dim intInspResultNo
Dim strPlantCd
Dim strInspClassCd
Dim IG1_import_group
Dim iErrorPosition
Dim LngRow

Dim i			'2003-03-01 Release 추가 
Dim SpdCount	'2003-03-01 Release 추가 
	
Redim IG1_import_group(5)	
Const Q270_IG1_I1_ief_supplied_select_char = 0
Const Q270_IG1_I2_q_inspection_details_insp_item_cd = 1
Const Q270_IG1_I2_q_inspection_details_insp_series = 2
Const Q270_IG1_I3_q_inspection_defect_type_defect_type_cd = 3
Const Q270_IG1_I4_q_inspection_defect_cause_defect_cause_cd = 4
Const Q270_IG1_I4_q_inspection_defect_cause_defect_qty = 5

Dim I2_q_inspection_result
Redim I2_q_inspection_result(2)			
Const Q270_I2_insp_result_no = 0
Const Q270_I2_plant_cd = 1
Const Q270_I2_insp_class_cd = 2	
' 전송된 데이타 받기 
LngMaxRow		= CInt(Request("txtMaxRows"))	
strInspReqNo	= Trim(Request("txtInspReqNo"))
intInspResultNo = 1
strPlantCd		= Trim(Request("txtPlantCd"))
strInspClassCd	= strinsp_class_cd
strUserId		= Request("txtInsrtUserId")
		
Dim txtSpread
txtSpread = Request("txtSpread")
SpdCount = CInt(Request("SpdCount"))	'2003-03-01 Release 추가 

For i = 1 to SpdCount
	txtSpread = txtSpread & Request("txtSpread" & i)	'2003-03-01 Release 추가 
Next

If txtSpread = "" Then					'2003-03-01 Release 추가 
	Response.End 
End If

Const C_SHEETMAXROWS_D = 100		
	
I2_q_inspection_result(Q270_I2_insp_result_no) = intInspResultNo
I2_q_inspection_result(Q270_I2_insp_class_cd) = strInspClassCd
I2_q_inspection_result(Q270_I2_plant_cd) = strPlantCd

Set PQIG170 = Server.CreateObject("PQIG170.cQmtInspDefCauseSvr")

lGrpCnt  = 0
If txtSpread <> "" Then
	arrRowVal = Split(txtSpread, gRowSep)
	Redim IG1_import_group(LngMaxRow,5)

	For LngRow = 1 To LngMaxRow
		arrColVal = Split(arrRowVal(LngRow-1), gColSep)
		lGrpCnt = lGrpCnt +1														'☜: Group Count
		strStatus = arrColVal(0)
		Select Case strStatus
			Case "C"			
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_item_cd) = arrColVal(1)
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_series) = arrColVal(2)	
				IG1_import_group(LngRow,Q270_IG1_I3_q_inspection_defect_type_defect_type_cd) = arrColVal(3)
				IG1_import_group(LngRow,Q270_IG1_I4_q_inspection_defect_cause_defect_cause_cd) = arrColVal(4)
				IG1_import_group(LngRow,Q270_IG1_I4_q_inspection_defect_cause_defect_qty) = UNIConvNum(arrColVal(5),0)
				IG1_import_group(LngRow,Q270_IG1_I1_ief_supplied_select_char) = "C"
			Case "U"	
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_item_cd) = arrColVal(1)
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_series) = arrColVal(2)	
				IG1_import_group(LngRow,Q270_IG1_I3_q_inspection_defect_type_defect_type_cd) = arrColVal(3)
				IG1_import_group(LngRow,Q270_IG1_I4_q_inspection_defect_cause_defect_cause_cd) = arrColVal(4)
				IG1_import_group(LngRow,Q270_IG1_I4_q_inspection_defect_cause_defect_qty) = UNIConvNum(arrColVal(5),0)
				IG1_import_group(LngRow,Q270_IG1_I1_ief_supplied_select_char) = "U"
			Case "D"
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_item_cd) = arrColVal(1)
				IG1_import_group(LngRow,Q270_IG1_I2_q_inspection_details_insp_series) = arrColVal(2)
				IG1_import_group(LngRow,Q270_IG1_I3_q_inspection_defect_type_defect_type_cd) = arrColVal(3)	
				IG1_import_group(LngRow,Q270_IG1_I4_q_inspection_defect_cause_defect_cause_cd) = arrColVal(4)
				IG1_import_group(LngRow,Q270_IG1_I1_ief_supplied_select_char) = "D"
	 	End Select			
	Next
						
	'/* 전체 삭제 관련 - START */					
	Call PQIG170.Q_MAINT_INSP_DEFT_CAUSE_SVR(gStrGlobalCollection, UCase(strInspReqNo), _
											I2_q_inspection_result, "N", IG1_import_group)
	'/* 전체 삭제 관련 - END */

	If CheckSYSTEMError(Err,True) = True Then
		Set PQIG170 = Nothing	
		Response.End
	End If
End If
	
Set PQIG170 = Nothing
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>
<%					
' Server Side 로직은 여기서 끝남 
'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
<Script Language=vbscript RUNAT=server>
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
</Script>
