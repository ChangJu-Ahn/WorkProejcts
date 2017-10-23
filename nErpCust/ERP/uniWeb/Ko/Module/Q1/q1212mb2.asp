<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")%>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1212MB1
'*  4. Program Name         : 기타검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG130
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
	
Dim PQBG130																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim lgIntFlgMode
Dim LngMaxRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt									'☜: Group Count
Dim strUserId
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspItemCd
Dim strRoutNo
Dim strOprNo
Dim IG1_import_group
Dim iErrorPosition
Dim LngRow

Const Q041_IG1_I1_q_wks_client_row_num = 0    
Const Q041_IG1_I2_q_inspection_condition_insp_series = 1
Const Q041_IG1_I2_q_inspection_condition_sample_qty = 2
Const Q041_IG1_I2_q_inspection_condition_accpt_decision_qty = 3
Const Q041_IG1_I2_q_inspection_condition_rejt_decision_qty = 4
Const Q041_IG1_I2_q_inspection_condition_accpt_decision_discreate = 5
Const Q041_IG1_I2_q_inspection_condition_max_defect_ratio = 6    
Const Q041_IG1_I3_ief_supplied_select_char = 7	
	
LngMaxRow		= CInt(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 
lgIntFlgMode	= CInt(Request("txtFlgMode"))					'☜: 저장시 Create/Update 판별 
strPlantCd		= UCase(Request("txtPlantCd"))
strItemCd		= UCase(Request("txtItemCd"))
strInspClassCd	= UCase(Request("cboInspClassCd"))
strInspItemCd	= UCase(Request("txtInspItemCd"))
strRoutNo		= UCase(Request("txtRoutNo"))
strOprNo		= UCase(Request("txtOprNo"))

Set PQBG130 = Server.CreateObject ("PQBG130.cQMtInspCndtnSvr")
	
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If	
	
lGrpCnt  = 0
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	Redim IG1_import_group(LngMaxRow,7)		
	For LngRow = 1 To LngMaxRow		    		
		arrColVal = Split(arrRowVal(LngRow-1), gColSep)		
		lGrpCnt = lGrpCnt +1														'☜: Group Count	
		strStatus = arrColVal(0)													'☜: Row 의 상태											
		Select Case strStatus

			Case "C"
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_insp_series) = arrColVal(1)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_sample_qty) = UNIConvNum(arrColVal(2),0)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_qty) = UNIConvNum(arrColVal(3),0)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_rejt_decision_qty) = UNIConvNum(arrColVal(4),0)					
				
				If arrColVal(5) <> "" Then
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_discreate) = UNIConvNum(arrColVal(5),0)
				Else
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_discreate) = ""
				End If
				
				If arrColVal(6) <> "" Then
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_max_defect_ratio) = UNIConvNum(arrColVal(6),0)
				Else
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_max_defect_ratio) = ""
				End If					
				
				IG1_import_group(LngRow,Q041_IG1_I1_q_wks_client_row_num) = arrColVal(7)	
				IG1_import_group(LngRow,Q041_IG1_I3_ief_supplied_select_char) =	"C"	
					
			Case "U"					
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_insp_series) = arrColVal(1)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_sample_qty) = UNIConvNum(arrColVal(2),0)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_qty) = UNIConvNum(arrColVal(3),0)
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_rejt_decision_qty) = UNIConvNum(arrColVal(4),0)					
				
				If arrColVal(5) <> "" Then
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_discreate) = UNIConvNum(arrColVal(5),0)
				Else
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_accpt_decision_discreate) = ""
				End If
				
				If arrColVal(6) <> "" Then
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_max_defect_ratio) = UNIConvNum(arrColVal(6),0)
				Else
					IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_max_defect_ratio) = ""
				End If					
				
				IG1_import_group(LngRow,Q041_IG1_I1_q_wks_client_row_num) = arrColVal(7)
				IG1_import_group(LngRow,Q041_IG1_I3_ief_supplied_select_char) =	"U"	
			Case "D"
				IG1_import_group(LngRow,Q041_IG1_I2_q_inspection_condition_insp_series) = arrColVal(1)
				IG1_import_group(LngRow,Q041_IG1_I1_q_wks_client_row_num) = arrColVal(2)
				IG1_import_group(LngRow,Q041_IG1_I3_ief_supplied_select_char) = "D"
	 	End Select
			
	Next
	
	Call PQBG130.Q_MT_INSP_CNDTN_SVR (gStrGlobalCollection, strInspClassCd, strInspItemCd, _
									UCase(strItemCd), strPlantCd, strRoutNo, _
									strOprNo, IG1_import_group, iErrorPosition)
	
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
		Response.End
	End If
	
End If
	
Set PQBG130 = Nothing                                                   '☜: Unload Comproxy
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
