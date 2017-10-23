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
'*  3. Program ID           : Q2413MB3
'*  4. Program Name         : 불량유형등록 
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

Const C_SHEETMAXROWS_D = 100

Dim strinsp_class_cd
strinsp_class_cd = "S"	'@@@주의 
Dim LngRow
'***** START
Dim LngMaxRow
'***** END
Dim intGroupCount 
Dim PQIG090
Dim lgStrPrevKeyM
Dim lglngHiddenRows
Dim lRow
Dim StrNextKey
Dim strInspReqNo
Dim strPlantCd
Dim strInspItemCd
Dim strInspSeries
Dim DefectTypeCd
Dim strData

Dim E1_q_inspection_result
Dim E2_b_plant
Dim E3_b_item
Dim E4_p_work_center
Dim E5_b_biz_partner
Dim E6_b_minor
Dim E7_b_minor
Dim E8_q_inspection_defect_type


Dim EG1_group_export
Dim EG2_group_export

Dim EG3_group_export
Redim EG3_group_export(1)
Const Q234_EG3_E1_q_inspection_defect_type_defect_type_cd = 0
Const Q234_EG3_E1_q_inspection_defect_type_defect_qty = 1
    
Dim EG4_group_export
Redim EG4_group_export(0)
Const Q235_EG4_E1_q_defect_type_defect_type_nm = 0

Dim E9_q_inspection_details
Redim E9_q_inspection_details(1)
Const Q234_E9_insp_item_cd = 0
Const Q234_E9_insp_series = 1

Dim I2_q_inspection_result
Redim I2_q_inspection_result(2)

Const Q234_I2_insp_result_no = 0
Const Q234_I2_plant_cd = 1
Const Q234_I2_insp_class_cd = 2

Dim I3_q_inspection_details
Redim I3_q_inspection_details(1)
Const Q234_I3_insp_item_cd = 0
Const Q234_I3_insp_series = 1
        


strPlantCd = Request("txtPlantCd")
strInspReqNo = Request("txtInspReqNo")
strInspItemCd = Request("txtInspItemCd")
strInspSeries = Request("txtInspSeries")

I2_q_inspection_result(Q234_I2_plant_cd)		= strPlantCd
I2_q_inspection_result(Q234_I2_insp_class_cd)	= strinsp_class_cd
I3_q_inspection_details(Q234_I3_insp_item_cd)	= strInspItemCd
I3_q_inspection_details(Q234_I3_insp_series)	= strInspSeries

lgStrPrevKeyM = Request("lgStrPrevKeyM")
lglngHiddenRows = CLng(Request("lglngHiddenRows"))
lRow = CLng(Request("lRow"))
'***** START
LngMaxRow = Request("txtMaxRows")
'***** END
If lgStrPrevKeyM <> "" Then
	DefectTypeCd = lgStrPrevKeyM
end if

Set PQIG090 = Server.CreateObject("PQIG090.cQListInspDefTypeSvr")


If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

call PQIG090.Q_LIST_INSP_DEFECT_TYPE_SVR  (gStrGlobalCollection, C_SHEETMAXROWS_D, strInspReqNo, _
										I2_q_inspection_result, I3_q_inspection_details , , DefectTypeCd , _
										E1_q_inspection_result, E2_b_plant, E3_b_item, E4_p_work_center, _
										E5_b_biz_partner, E6_b_minor, E7_b_minor, EG1_group_export, _
										EG2_group_export, EG3_group_export, EG4_group_export, _
										E8_q_inspection_defect_type, E9_q_inspection_details)

'-----------------------
'Com Action Area
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG090 = Nothing	
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG090 = Nothing

strData = ""
intGroupCount = UBound(EG3_group_export, 1)
If intGroupCount < C_SHEETMAXROWS_D Then
	intGroupCount = UBound(EG3_group_export, 1) + 1
End if

Dim i
For i = 0 To UBound(EG3_group_export,1)
	If i < C_SHEETMAXROWS_D Then
		
			strData = strData & Chr(11) & Trim(ConvSPChars(EG3_group_export(i, Q234_EG3_E1_q_inspection_defect_type_defect_type_cd)))				
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ConvSPChars(EG4_group_export(i,Q235_EG4_E1_q_defect_type_defect_type_nm))
			strData = strData & Chr(11) & UniNumClientFormat(EG3_group_export(i,Q234_EG3_E1_q_inspection_defect_type_defect_qty), ggQty.DecPoint ,0)
			'***** START
			strData = strData & Chr(11) & lRow 							'Parent Row No
			strData = strData & Chr(11) & lglngHiddenRows + i + 1 		'Flag
			strData = strData & Chr(11) & LngMaxRow + i + 1
       		'***** END
      		strData = strData & Chr(11) & Chr(12)
	Else		
		StrNextKey = EG3_group_export(i, 0)		
	End if
Next
%>
<Script Language="vbscript"> 
    With Parent
		.ggoSpread.Source = .frm1.vspdData2
		.frm1.vspdData2.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=strData%>"
	
		.lgStrPrevKeyM(<%=lRow-1%>) = "<%=ConvSPChars(StrNextKey)%>"
		.lglngHiddenRows(<%=lRow-1%>) = "<%=lglngHiddenRows + intGroupCount%>"
		
		If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKeyM(<%=lRow-1%>) <> "" Then
			call .DbQuery2(<%=lRow%>, True)
		Else		
			<% ' Request값을 hidden input으로 넘겨줌 %>
			.frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
			.frm1.hInspItemCd.value = "<%=ConvSPChars(strInspItemCd)%>"
			.frm1.hInspSeries.value = "<%=strInspSeries%>"
			Call .DbQueryOk2("<%=intGroupCount%>")
	    End If		
	End with
</Script>
<%
Set PQIG090 = Nothing  
%>