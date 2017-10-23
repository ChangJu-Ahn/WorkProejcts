<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1312MB2
'*  4. Program Name         : 불량원인 정보등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG240
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

Dim PQBG240																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim lGrpCnt								'☜: Group Count
Dim strUserId
Dim strPlantCd
Dim strInspClassCd
Dim strDefectTypeCd
Dim IG1_import_group
Dim iErrorPosition
	
Const Q082_IG1_I1_q_wks_client_row_num = 0    
Const Q082_IG1_I2_ief_supplied_select_char = 1    
Const Q082_IG1_I3_q_defect_cause_defect_cause_cd = 2
Const Q082_IG1_I3_q_defect_cause_defect_cause_nm = 3
    
Const C_SHEETMAXROWS_D = 100
	
LngMaxRow = CInt(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 
lgIntFlgMode = CInt(Request("txtFlgMode"))					'☜: 저장시 Create/Update 판별 
strPlantCd = Request("txtPlantCd")
strInspClassCd = Request("cboInspClassCd")
	
Set PQBG240 = Server.CreateObject("PQBG240.cQMtDefCauseSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

lGrpCnt  = 0
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	Redim IG1_import_group(LngMaxRow,3)
			
	For LngRow = 1 To LngMaxRow		    		
		arrColVal = Split(arrRowVal(LngRow-1), gColSep)		
		lGrpCnt = lGrpCnt +1													'☜: Group Count	
		strStatus = arrColVal(0)
																			'☜: Row 의 상태 
		Select Case strStatus
			Case "C"
				IG1_import_group(LngRow,Q082_IG1_I3_q_defect_cause_defect_cause_cd) = arrColVal(1)					
				IG1_import_group(LngRow,Q082_IG1_I3_q_defect_cause_defect_cause_nm) = arrColVal(2)					
				IG1_import_group(LngRow,Q082_IG1_I1_q_wks_client_row_num) = arrColVal(3)					
				IG1_import_group(LngRow,Q082_IG1_I2_ief_supplied_select_char) = "C"									
			Case "U"					
				IG1_import_group(LngRow,Q082_IG1_I3_q_defect_cause_defect_cause_cd) = arrColVal(1)					
				IG1_import_group(LngRow,Q082_IG1_I3_q_defect_cause_defect_cause_nm) = arrColVal(2)					
				IG1_import_group(LngRow,Q082_IG1_I1_q_wks_client_row_num) = arrColVal(3)					
				IG1_import_group(LngRow,Q082_IG1_I2_ief_supplied_select_char) = "U"									
			Case "D"
				IG1_import_group(LngRow,Q082_IG1_I3_q_defect_cause_defect_cause_cd) = arrColVal(1)		
				IG1_import_group(LngRow,Q082_IG1_I1_q_wks_client_row_num) = arrColVal(2)	
				IG1_import_group(LngRow,Q082_IG1_I2_ief_supplied_select_char) = "D"
		End Select				
	Next
		
	Call PQBG240.Q_MAINT_DEFECT_CAUSE_SVR (gStrGlobalCollection, strInspClassCd, _
										strPlantCd, IG1_import_group, iErrorPosition)

	If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
		Set PQBG240 = Nothing  
		Response.End			
	End If		
			
End If	
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