<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1111MB1
'*  4. Program Name         : 측정기 정보등록 
'*  5. Program Desc         : 측정기 정보등록 
'*  6. Component List       : PQBG020
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
 
Const C_SHEETMAXROWS_D = 100

Dim PQBG020
Dim EG1_export_group
Dim I1_q_measurement_equipment_measmt_equipmt_cd
Dim E1_q_measurement_equipment_measmt_equipmt_cd

Const Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_cd = 0
Const Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_nm = 1

Dim strMeasmtEquipmtCd

Dim StrNextKey		' 다음 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim lgStrPrevKey
Dim strData

	LngMaxRow = Request("txtMaxRows")
	strMeasmtEquipmtCd = Request("txtMeasmtEquipmtCd")
	lgStrPrevKey = Request("lgStrPrevKey")


	If lgStrPrevKey = "" then
		I1_q_measurement_equipment_measmt_equipmt_cd = strMeasmtEquipmtCd
	Else
		I1_q_measurement_equipment_measmt_equipmt_cd = lgStrPrevKey
	End If


	Set PQBG020 = Server.CreateObject("PQBG020.cQListMeaEquSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End if

	Call PQBG020.Q_LIST_MEA_EQU_SVR (gStrGlobalCollection, _
									 C_SHEETMAXROWS_D, _
									 I1_q_measurement_equipment_measmt_equipmt_cd, _
									 EG1_export_group)
		        
	If CheckSYSTEMError(Err,True) Then
		Set PQBG020 = Nothing
		Response.End
	End If

	Dim iTotalStr
	Dim TmpBuffer
	ReDim TmpBuffer(UBound(EG1_export_group))

	For LngRow = 0 to UBound(EG1_export_group)
	
		If LngRow < C_SHEETMAXROWS_D Then
			
			strData = chr(11) & ConvSPChars(EG1_export_group(LngRow,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_cd)) & _
					  chr(11) & ConvSPChars(EG1_export_group(LngRow,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_nm)) & _
					  chr(11) & LngMaxRow + LngRow + 1 & Chr(11) & Chr(12)
			
			TmpBuffer(LngRow) = strData
		Else
			StrNextKey = EG1_export_group(LngRow,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_cd)
		End If
	Next

iTotalStr = Join(TmpBuffer, "")

%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	' Request값을 hidden input으로 넘겨줌 
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hMeasmtEquipmtCd.Value = "<%=ConvSPChars(strNextKey)%>"
		.DbQueryOk
    End If
	    
    If UCase(Trim(.frm1.txtMeasmtEquipmtCd.Value)) = "<%=ConvSPChars(EG1_export_group(0,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_cd))%>" Then
		.frm1.txtMeasmtEquipmtCd.Value = "<%=ConvSPChars(EG1_export_group(0,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_cd))%>"
		.frm1.txtMeasmtEquipmtNm.Value = "<%=ConvSPChars(EG1_export_group(0,Q005_EG1_E1_q_measurement_equipment_measmt_equipmt_nm))%>"
	Else
		.frm1.txtMeasmtEquipmtNm.Value = ""
	End If
End with
</Script>
<%
Set PQBG020 = Nothing
%>