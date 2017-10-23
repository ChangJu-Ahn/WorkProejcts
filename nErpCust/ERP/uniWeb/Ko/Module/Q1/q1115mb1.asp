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
'*  3. Program ID           : Q1115MB1
'*  4. Program Name         : 연/월 품질목표등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG100
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

Dim PQBG100

Dim EG1_export_group

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim i
	          
Dim I1_plant_cd, I2_insp_class_cd, I3_year

Dim E1_q_yearly_target
	Const Q031_E1_insp_class_cd = 0
    Const Q031_E1_yr = 1
    Const Q031_E1_target_value = 2
    Const Q031_E1_plant_cd = 3
    Const Q031_E1_plant_nm = 4
    Const Q031_E1_defect_ratio_unit_cd = 5
    Const Q031_E1_defect_ratio_unit_nm = 6

Dim EG1_q_monthly_target
'Const Q031_EG1_mnth = 0
'Const Q031_EG1_monthly_target_value = 1

	
	
	I1_plant_cd = Request("txtPlantCd")
	I2_insp_class_cd = Request("cboInspClassCd")
	I3_year = Request("txtYr")

	Set PQBG100 = Server.CreateObject("PQBG100.cQListTargetSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End if

	Call PQBG100.Q_LIST_TARGET_SVR (gStrGlobalCollection, _
									I1_plant_cd, _
									I2_insp_class_cd, _
									I3_year, _
									E1_q_yearly_target, _
									EG1_q_monthly_target)

'##############################################################################
	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "125000"	'공장이 존재하지 않습니다.
			If CheckSYSTEMError(Err,True) = True Then
				%>
				<Script Language=vbscript>			
					Parent.frm1.txtPlantNm1.Value = ""
					Parent.frm1.txtPlantCd1.Focus
				</Script>
				<%
				Set PQBG100 = Nothing
				Response.End
			End If
		Case Else
			If CheckSYSTEMError(Err,True) = True Then
				Set PQBG100 = Nothing
				Response.End
			End If	
	End Select
'##############################################################################
	Set PQBG100 = Nothing

    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent.frm1 " & vbCr
	Response.Write "	.txtPlantNm1.Value	= """ & ConvSPChars(E1_q_yearly_target(Q031_E1_plant_nm)) & """" & vbCr
	Response.Write "	.txtPlantCd2.Value  = """ & ConvSPChars(E1_q_yearly_target(Q031_E1_plant_cd)) & """" & vbCr
	Response.Write "	.txtPlantNm2.value  = """ & ConvSPChars(E1_q_yearly_target(Q031_E1_plant_nm)) & """" & vbCr
	Response.Write "	.txtYr2.Text		= """ & E1_q_yearly_target(Q031_E1_yr) & """" & vbCr
	Response.Write "	.cboInspClassCd2.value      = """ & E1_q_yearly_target(Q031_E1_insp_class_cd) & """" & vbCr
	Response.Write "	.cboDefectRatioUnitCd.value = """ & ConvSPChars(E1_q_yearly_target(Q031_E1_defect_ratio_unit_cd)) & """" & vbCr
	Response.Write "	.txtYrTargetValue.Text      = """ & UniNumClientFormat(E1_q_yearly_target(Q031_E1_target_value), 2, 0) & """" & vbCr
	
	For i = 0 To Ubound(EG1_q_monthly_target, 1)
	
		Select Case EG1_q_monthly_target(i,0)
			Case "01"
				Response.Write "	.txtMnthTargetValue1.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "02"
				Response.Write "	.txtMnthTargetValue2.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "03"
				Response.Write "	.txtMnthTargetValue3.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "04"
				Response.Write "	.txtMnthTargetValue4.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "05"
				Response.Write "	.txtMnthTargetValue5.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "06"
				Response.Write "	.txtMnthTargetValue6.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "07"
				Response.Write "	.txtMnthTargetValue7.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "08"
				Response.Write "	.txtMnthTargetValue8.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "09"
				Response.Write "	.txtMnthTargetValue9.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "10"
				Response.Write "	.txtMnthTargetValue10.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "11"
				Response.Write "	.txtMnthTargetValue11.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
			Case "12"
				Response.Write "	.txtMnthTargetValue12.Text = """ & UniNumClientFormat(EG1_q_monthly_target(i,1), 2 ,0) & """" & vbCr
		End Select
	Next
	Response.Write "End with " & vbcr
	
	Response.Write "	lgNextNo				= """" " & vbCr
	Response.Write "	lgPrevNo				= """" " & vbCr
    Response.Write "parent.DbQueryOK	"	& vbCr
	Response.Write "</Script>	" & vbCr
	
	Response.End

%>

