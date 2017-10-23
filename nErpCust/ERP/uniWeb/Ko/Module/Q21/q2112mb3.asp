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
'*  3. Program ID           : Q2112MB3
'*  4. Program Name         : 내역등록 
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
strinsp_class_cd = "R"	'@@@주의 
Dim LngRow
'****** START
Dim LngMaxRow
'****** END
Dim intGroupCount 
Dim PQIG070
Dim lgStrPrevKeyM
Dim lglngHiddenRows
Dim lRow
Dim StrNextKey
Dim strInspReqNo
Dim strPlantCd
Dim strInspItemCd
Dim strInspSeries
Dim strData
Dim i

Dim E1_q_inspection_result
Dim E2_b_plant
Dim E3_b_item
Dim E4_p_work_center
Dim E5_b_biz_partner
Dim E6_b_minor
Dim E7_b_minor
Dim EG1_group_export
	    
Dim I1_q_inspection_measured_values_sample_no

Dim I2_q_inspection_details
Const Q227_I2_insp_item_cd = 0
Const Q227_I2_insp_series = 1
ReDim I2_q_inspection_details(1)

Dim I3_q_inspection_result
Const Q227_I3_insp_result_no = 0
Const Q227_I3_plant_cd = 1
Const Q227_I3_insp_class_cd = 2
ReDim I3_q_inspection_result(2)

Dim I4_q_inspection_request_insp_req_no
	
Const Q227_E1_insp_result_no = 0
Const Q227_E1_insp_class_cd = 1
Const Q227_E1_insp_dt = 2
Const Q227_E1_insp_qty = 3
Const Q227_E1_defect_qty = 4
Const Q227_E1_decision = 5
Const Q227_E1_inspector_cd = 6
Const Q227_E1_rmk = 7
Const Q227_E1_bp_cd = 8
Const Q227_E1_wc_cd = 9
Const Q227_E1_item_cd = 10
Const Q227_E1_plant_cd = 11
Const Q227_E1_lot_no = 12
Const Q227_E1_lot_sub_no = 13
Const Q227_E1_lot_size = 14
Const Q227_E1_sl_cd = 15
Const Q227_E1_sl_cd_for_good = 16
Const Q227_E1_sl_cd_for_defect = 17
Const Q227_E1_status_flag = 18
Const Q227_E1_transfer_flag = 19

Const Q227_E2_plant_nm = 0

Const Q227_E3_item_nm = 0
Const Q227_E3_spec = 1
Const Q227_E3_basic_unit = 2

Const Q227_E4_wc_nm = 0
Const Q227_E5_bp_nm = 0
Const Q227_E6_minor_nm = 0
Const Q227_E7_minor_nm = 0

Const Q227_EG1_E1_sample_no = 0
Const Q227_EG1_E1_insp_class_cd = 1
Const Q227_EG1_E1_meas_value = 2
Const Q227_EG1_E1_defect_flag = 3

strInspReqNo = Request("txtInspReqNo")
strPlantCd = Request("txtPlantCd")	
strInspItemCd = Request("txtInspItemCd")
strInspSeries = Request("txtInspSeries")
lgStrPrevKeyM = Request("lgStrPrevKeyM")
lglngHiddenRows = CLng(Request("lglngHiddenRows"))
lRow = CLng(Request("lRow"))
LngMaxRow = Request("txtMaxRows")
	
Set PQIG070 = Server.CreateObject("PQIG070.cQListInspMeaValSvr")

If CheckSystemError(Err,True) Then
	Response.End					
End If

I1_q_inspection_measured_values_sample_no = UniConvNum(lgStrPrevKeyM, 0)

I2_q_inspection_details(Q227_I2_insp_item_cd) = strInspItemCd
I2_q_inspection_details(Q227_I2_insp_series) = UniConvNum(strInspSeries, 0)

I3_q_inspection_result(Q227_I3_insp_result_no) = 1
I3_q_inspection_result(Q227_I3_plant_cd) = strPlantCd
I3_q_inspection_result(Q227_I3_insp_class_cd) = strinsp_class_cd	'@@@주의 

I4_q_inspection_request_insp_req_no = strInspReqNo

Call PQIG070.Q_LIST_INSP_MEAS_VALUE_SVR(gStrGlobalCollection, _
										C_SHEETMAXROWS_D, _
										I1_q_inspection_measured_values_sample_no, _
										I2_q_inspection_details, _
										I3_q_inspection_result, _
										I4_q_inspection_request_insp_req_no, _
										E1_q_inspection_result, _
										E2_b_plant, _
										E3_b_item, _
										E4_p_work_center, _
										E5_b_biz_partner, _
										E6_b_minor, _
										E7_b_minor, _
										EG1_group_export)


If CheckSystemError(Err,True) Then
	Set PQIG070 = Nothing	
%>
<Script Language="vbscript">   
	Parent.frm1.cmdInsertSampleRows.Disabled = false
</Script>    
<%
	Response.End				
End If
	
'Export To Spread2
strData = ""
'**** hidden NextKey
intGroupCount = UBound(EG1_group_export, 1)
If intGroupCount < C_SHEETMAXROWS_D Then
	intGroupCount = UBound(EG1_group_export, 1) + 1
End If

For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q227_EG1_E1_sample_no)))

		If Trim(ConvSPChars(EG1_group_export(i, Q227_EG1_E1_meas_value))) = "" Then
			strData = strData & Chr(11) & ""
		Else
			strData = strData & Chr(11) & UniNumClientFormat(Trim(ConvSPChars(EG1_group_export(i, Q227_EG1_E1_meas_value))), 4, 0)
		End If
				
		If Trim(ConvSPChars(EG1_group_export(i, Q227_EG1_E1_defect_flag))) = "G" Then
			strData = strData & Chr(11) & "0"
		Else
			strData = strData & Chr(11) & "1"
		End If		
		'****** START
		strData = strData & Chr(11) & lRow 							'Parent Row No
		strData = strData & Chr(11) & lglngHiddenRows + i + 1 		'Flag
		strData = strData & Chr(11) & LngMaxRow + i + 1 
		'****** END
		strData = strData & Chr(11) & Chr(12)
    Else
		StrNextKey = EG1_group_export(i, Q227_EG1_E1_sample_no)
    End If
Next
%>
<Script Language="vbscript">   
Dim StrData
With Parent
	.ggoSpread.Source = .frm1.vspdData2 
	.frm1.vspdData2.Redraw = False
	.ggoSpread.SSShowDataByClip "<%=strData%>"	
	.lgStrPrevKeyM(<%=lRow-1%>) = "<%=StrNextKey%>"
	.lglngHiddenRows(<%=lRow-1%>) = "<%=lglngHiddenRows + intGroupCount%>"
		
	If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKeyM(<%=lRow-1%>) <> "" Then
		Call .DbQuery2(<%=lRow%>, True)
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