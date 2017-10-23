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
'*  3. Program ID           : Q1115MB2
'*  4. Program Name         : 연/월 품질목표등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG090
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/12
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
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

Dim PQBG090																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim lgIntFlgMode
	
Dim dblMmValue
Dim i
Dim iErrorPosition
Dim iCommandSent
	
Dim I1_q_yearly_target
    'Const Q025_I1_plant_cd = 0
    'Const Q025_I1_insp_class_cd = 1
    'Const Q025_I1_yr = 2
    'Const Q025_I1_target_value = 3
    'Const Q025_I1_defect_unit_cd = 4
ReDim I1_q_yearly_target(4)

Dim IG1_q_monthly_target
    'Const Q025_IG1_row_num = 0
    'Const Q025_IG1_mnth = 1
    'Const Q025_IG1_monthly_target_value = 2
ReDim IG1_q_monthly_target(11, 2)

	lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent  = "CREATE"							
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent  = "UPDATE"					
	End If

	I1_q_yearly_target(0) = UCase(Trim(Request("txtPlantCd2")))
	I1_q_yearly_target(1) = Request("cboInspClassCd2")	
	I1_q_yearly_target(2) = Request("txtYr2")	
	I1_q_yearly_target(3) = UNIConvNum(Request("txtYrTargetValue"), 0)
	I1_q_yearly_target(4) = Request("cboDefectRatioUnitCd")

	dblMmValue = Array(UNIConvNum(Request("txtMnthTargetValue1"), 0), UNIConvNum(Request("txtMnthTargetValue2"), 0), _
					   UNIConvNum(Request("txtMnthTargetValue3"), 0), UNIConvNum(Request("txtMnthTargetValue4"), 0), _
					   UNIConvNum(Request("txtMnthTargetValue5"), 0), UNIConvNum(Request("txtMnthTargetValue6"), 0), _
					   UNIConvNum(Request("txtMnthTargetValue7"), 0), UNIConvNum(Request("txtMnthTargetValue8"), 0), _
					   UNIConvNum(Request("txtMnthTargetValue9"), 0), UNIConvNum(Request("txtMnthTargetValue10"), 0), _
					   UNIConvNum(Request("txtMnthTargetValue11"), 0), UNIConvNum(Request("txtMnthTargetValue12"), 0))
	For i = 0 to 11
		IG1_q_monthly_target(i,0) = i + 1
		IG1_q_monthly_target(i,1) = Right("0" & Cstr(i + 1),2)
		IG1_q_monthly_target(i,2) = dblMmValue(i)
	Next		

	Set PQBG090 = Server.CreateObject("PQBG090.cQMaintTargetSvr")

	If CheckSYSTEMError(Err,True) Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End
	End If

	Call PQBG090.Q_MAINT_TARGET_SVR (gStrGlobalCollection, _
									iCommandSent, _
									I1_q_yearly_target, _
									IG1_q_monthly_target, _
									iErrorPosition)

'##############################################################################
	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "125000"	'공장이 존재하지 않습니다.
			If CheckSYSTEMError(Err,True) = True Then
				%>
				<Script Language=vbscript>			
					Parent.frm1.txtPlantNm2.Value = ""
					Parent.frm1.txtPlantCd2.Focus
				</Script>
				<%
				Set PQBG090 = Nothing
				Response.End
			End If
		Case Else
			If CheckSYSTEMError(Err,True) = True Then
				Call MonthFocus(iErrorPosition, I_MKSCRIPT)
				Set PQBG090 = Nothing  
				Response.End
			End If	
	End Select
'##############################################################################

	Set PQBG090 = Nothing
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>

<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : MonthFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function MonthFocus(Byval lRow, Byval iLoc)
	Dim strHTML
	If iLoc = I_INSCRIPT Then
		Select Case lRow
			Case 1
				strHTML = "parent.frm1.txtMnthTargetValue1.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue1.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue1.SelLength = len(parent.frm1.txtMnthTargetValue1.Text) " & vbCrLf
			Case 2
				strHTML = "parent.frm1.txtMnthTargetValue2.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue2.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue2.SelLength = len(parent.frm1.txtMnthTargetValue2.Text) " & vbCrLf
			Case 3
				strHTML = "parent.frm1.txtMnthTargetValue3.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue3.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue3.SelLength = len(parent.frm1.txtMnthTargetValue3.Text) " & vbCrLf
			Case 4
				strHTML = "parent.frm1.txtMnthTargetValue4.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue4.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue4.SelLength = len(parent.frm1.txtMnthTargetValue4.Text) " & vbCrLf
			Case 5
				strHTML = "parent.frm1.txtMnthTargetValue5.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue5.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue5.SelLength = len(parent.frm1.txtMnthTargetValue5.Text) " & vbCrLf
			Case 6
				strHTML = "parent.frm1.txtMnthTargetValue6.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue6.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue6.SelLength = len(parent.frm1.txtMnthTargetValue6.Text) " & vbCrLf
			Case 7
				strHTML = "parent.frm1.txtMnthTargetValue7.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue7.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue7.SelLength = len(parent.frm1.txtMnthTargetValue7.Text) " & vbCrLf
			Case 8
				strHTML = "parent.frm1.txtMnthTargetValue8.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue8.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue8.SelLength = len(parent.frm1.txtMnthTargetValue8.Text) " & vbCrLf
			Case 9
				strHTML = "parent.frm1.txtMnthTargetValue9.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue9.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue9.SelLength = len(parent.frm1.txtMnthTargetValue9.Text) " & vbCrLf
			Case 10
				strHTML = "parent.frm1.txtMnthTargetValue10.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue10.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue10.SelLength = len(parent.frm1.txtMnthTargetValue10.Text) " & vbCrLf
			Case 11
				strHTML = "parent.frm1.txtMnthTargetValue11.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue11.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue11.SelLength = len(parent.frm1.txtMnthTargetValue11.Text) " & vbCrLf
			Case 12
				strHTML = "parent.frm1.txtMnthTargetValue12.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue12.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue12.SelLength = len(parent.frm1.txtMnthTargetValue12.Text) " & vbCrLf
		End Select
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		Select Case lRow
			Case 1
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue1.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue1.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue1.SelLength = len(parent.frm1.txtMnthTargetValue1.Text) " & vbCrLf
			Case 2
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue2.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue2.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue2.SelLength = len(parent.frm1.txtMnthTargetValue2.Text) " & vbCrLf
			Case 3
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue3.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue3.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue3.SelLength = len(parent.frm1.txtMnthTargetValue3.Text) " & vbCrLf
			Case 4
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue4.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue4.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue4.SelLength = len(parent.frm1.txtMnthTargetValue4.Text) " & vbCrLf
			Case 5
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue5.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue5.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue5.SelLength = len(parent.frm1.txtMnthTargetValue5.Text) " & vbCrLf
			Case 6
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue6.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue6.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue6.SelLength = len(parent.frm1.txtMnthTargetValue6.Text) " & vbCrLf
			Case 7
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue7.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue7.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue7.SelLength = len(parent.frm1.txtMnthTargetValue7.Text) " & vbCrLf
			Case 8
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue8.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue8.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue8.SelLength = len(parent.frm1.txtMnthTargetValue8.Text) " & vbCrLf
			Case 9
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue9.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue9.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue9.SelLength = len(parent.frm1.txtMnthTargetValue9.Text) " & vbCrLf
			Case 10
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue10.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue10.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue10.SelLength = len(parent.frm1.txtMnthTargetValue10.Text) " & vbCrLf
			Case 11
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue11.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue11.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue11.SelLength = len(parent.frm1.txtMnthTargetValue11.Text) " & vbCrLf
			Case 12
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue12.focus" & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue12.SelStart = 0 " & vbCrLf
				strHTML = strHTML & "parent.frm1.txtMnthTargetValue12.SelLength = len(parent.frm1.txtMnthTargetValue12.Text) " & vbCrLf
		End Select
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>
