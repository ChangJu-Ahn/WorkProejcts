<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1114MB1
'*  4. Program Name         : 검사분류별 불량률단위등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG080
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

Dim PQBG080													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim strPlant
Dim strData
Dim E1_b_plant
Dim EG1_group_export
	
	Const C_SHEETMAXROWS_D = 100
    
    Const Q024_E1_plant_cd = 0
    Const Q024_E1_plant_nm = 1

    Const Q024_EG1_minor_nm = 0
    Const Q024_EG1_defect_ratio_unit_cd = 1
    Const Q024_EG1_insp_class_cd = 2
    	   
	lgStrPrevKey = Request("lgStrPrevKey")
	LngMaxRow = Request("txtMaxRows")
	strPlant = Trim(Request("txtPlantCd"))

	Set PQBG080 = Server.CreateObject("PQBG080.cQListDefRtInsClsSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call PQBG080.Q_LIST_DFCT_RTO_INS_CLS_SVR(gStrGlobalCollection, _
											 C_SHEETMAXROWS_D, _
											 lgStrPrevKey, _
											 strPlant, _
											 E1_b_plant, _
											 EG1_group_export)

'##############################################################################
	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "125000"	'공장이 존재하지 않습니다.
			If CheckSYSTEMError(Err,True) = True Then
				%>
				<Script Language=vbscript>			
					Parent.frm1.txtPlantNm.Value = ""
					Parent.frm1.txtPlantCd.Focus
					Parent.frm1.hPlantCd.value = ""
				</Script>
				<%
				Set PQBG080 = Nothing
				Response.End
			End If
		Case Else
			If CheckSYSTEMError(Err,True) = True Then
				%>
				<Script Language=vbscript>			
					Parent.frm1.hPlantCd.value = ""
				</Script>
				<%

				Set PQBG080 = Nothing
				Response.End
			End If	
	End Select
'##############################################################################

	Dim TmpBuffer
	Dim iTotalStr
	ReDim TmpBuffer(UBound(EG1_group_export, 1))
		  
	For LngRow = 0 To UBound(EG1_group_export, 1)
	    If LngRow < C_SHEETMAXROWS_D Then
			
			strData = Chr(11) & Trim(ConvSPChars(EG1_group_export(LngRow, Q024_EG1_minor_nm))) & _
					  Chr(11) & ConvSPChars(EG1_group_export(LngRow, Q024_EG1_defect_ratio_unit_cd)) & _
					  Chr(11) & Trim(ConvSPChars(EG1_group_export(LngRow, Q024_EG1_insp_class_cd))) & _
					  Chr(11) & LngMaxRow + LngRow + 1 & _
					  Chr(11) & Chr(12)
			TmpBuffer(LngRow) = strData
	    Else
			StrNextKey = EG1_group_export(LngRow,Q024_EG1_defect_ratio_unit_cd)
			
	    End If
	Next  

	iTotalStr = Join(TmpBuffer, "")

	Set PQBG080 = Nothing
%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"		
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlant)%>"
		.DbQueryOk
	End If
	
    If UCase(Trim(.frm1.txtPlantCd.Value)) = "<%=ConvSPChars(E1_b_plant(Q024_E1_plant_cd))%>" Then
		.frm1.txtPlantCd.Value = "<%=ConvSPChars(E1_b_plant(Q024_E1_plant_cd))%>"
		.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_b_plant(Q024_E1_plant_nm))%>"
	Else
		.frm1.txtPlantNm.Value = ""
	End If

End with			
</Script>
