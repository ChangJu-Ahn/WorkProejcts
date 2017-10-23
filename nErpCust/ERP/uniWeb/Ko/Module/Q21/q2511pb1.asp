<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2511PB1
'*  4. Program Name         : 
'*  5. Program Desc         : 검사의뢰현황 팝업 
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


Dim PQIG040

Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim StrNextKey
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspReqNo
Dim strBpCd
Dim strCustCd
Dim strWcCd
Dim strFrInspReqDt
Dim strToInspReqDt
Dim lgStrPrevKey
Dim StrData
Dim i
Dim EG1_export_group
Dim strInspReqDtFr
Dim strInspReqDtTo


Const C_SHEETMAXROWS_D = 100

Dim I5_q_inspection_request
ReDim I5_q_inspection_request(4)
'[CONVERSION INFORMATION]  IMPORTS View 상수 
Const Q217_I5_insp_req_no = 0    '[CONVERSION INFORMATION]  View Name : import q_inspection_request
Const Q217_I5_insp_class_cd = 1
Const Q217_I5_bp_cd = 2
Const Q217_I5_wc_cd = 3
Const Q217_I5_insp_status = 4

Dim E1_b_plant
ReDim E1_b_plant(1)

'[CONVERSION INFORMATION]  EXPORTS View 상수 
Const Q217_E1_plant_cd = 0    '[CONVERSION INFORMATION]  View Name : export_header b_plant
Const Q217_E1_plant_nm = 1

'[CONVERSION INFORMATION]  EXPORTS Group View 상수 
Const Q217_EG1_E1_plant_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_plant
Const Q217_EG1_E1_plant_nm = 1
Const Q217_EG1_E2_insp_req_no = 2    '[CONVERSION INFORMATION]  View Name : export q_inspection_request
Const Q217_EG1_E2_insp_class_cd = 3
Const Q217_EG1_E2_insp_req_dt = 4
Const Q217_EG1_E2_bp_cd = 5
Const Q217_EG1_E2_wc_cd = 6
Const Q217_EG1_E2_por_no = 7
Const Q217_EG1_E2_por_seq = 8
Const Q217_EG1_E2_prodt_no = 9
Const Q217_EG1_E2_lot_no = 10
Const Q217_EG1_E2_lot_sub_no = 11
Const Q217_EG1_E2_lot_size = 12
Const Q217_EG1_E2_tracking_no = 13
Const Q217_EG1_E2_insp_status = 14
Const Q217_EG1_E2_accum_lot_size = 15
Const Q217_EG1_E2_document_no = 16
Const Q217_EG1_E2_document_seq_no = 17
Const Q217_EG1_E2_report_seq = 18
Const Q217_EG1_E3_item_cd = 19    '[CONVERSION INFORMATION]  View Name : export b_item
Const Q217_EG1_E3_item_nm = 20
Const Q217_EG1_E4_wc_nm = 21    '[CONVERSION INFORMATION]  View Name : export p_work_center
Const Q217_EG1_E5_bp_nm = 22    '[CONVERSION INFORMATION]  View Name : export b_biz_partner
Const Q217_EG1_E6_minor_nm = 23    '[CONVERSION INFORMATION]  View Name : export_nm_for_insp_status b_minor
		
strPlantCd = Request("txtPlantCd")
strItemCd = Request("txtItemCd")
strInspClassCd = Request("txtInspClassCd")
strInspReqNo = Request("txtInspReqNo")
strBpCd = Request("txtBpCd")
strCustCd = Request("txtCustCd")
strWcCd = Request("txtWcCd")
strFrInspReqDt = Request("txtFrInspReqDt")
strToInspReqDt = Request("txtToInspReqDt")
lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")

If strFrInspReqDt <> "" then
	strInspReqDtFr = UNIConvDate(strFrInspReqDt)
End If

If strToInspReqDt <> "" then
	strInspReqDtTo = UNIConvDate(strToInspReqDt)
End If

If lgStrPrevKey = "" Then
	I5_q_inspection_request(Q217_I5_insp_req_no) = strInspReqNo
Else
	I5_q_inspection_request(Q217_I5_insp_req_no) = lgStrPrevKey
End If

I5_q_inspection_request(Q217_I5_insp_class_cd) = strInspClassCd
I5_q_inspection_request(Q217_I5_bp_cd) = ""
I5_q_inspection_request(Q217_I5_wc_cd) = ""
I5_q_inspection_request(Q217_I5_insp_status) = "N"

Select Case strInspClassCd
	Case "R"
		I5_q_inspection_request(Q217_I5_bp_cd) = strBpCd	
	Case "P"
		I5_q_inspection_request(Q217_I5_wc_cd) = strWcCd	
	Case "F"
	
	Case "S"
		I5_q_inspection_request(Q217_I5_bp_cd) = strCustCd
End Select 

Set PQIG040 = Server.CreateObject("PQIG040.cQListInspReqSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG040.Q_LIST_INSP_REQUEST_SVR(gStrGlobalCollection, _
									 C_SHEETMAXROWS_D, _
									 strPlantCd, _
									 strItemCd, _
									 strInspReqDtFr, _
									 strInspReqDtTo, _
									 I5_q_inspection_request, _
									 E1_b_plant, _
									 EG1_export_group)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG040 = Nothing
	Response.End
End If

For i = 0 To UBound(EG1_export_group, 1)
    If i < C_SHEETMAXROWS_D Then
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_insp_req_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E3_item_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E3_item_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_bp_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E5_bp_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_wc_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E4_wc_nm)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_insp_req_dt))))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_lot_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_lot_sub_no)))
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(i, Q217_EG1_E2_lot_size), ggQty.DecPoint ,0)
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_por_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_por_seq)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_document_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q217_EG1_E2_document_seq_no)))
		strData = strData & Chr(11) & LngMaxRow + i + 1
		strData = strData & Chr(11) & Chr(12)
    Else
		StrNextKey = EG1_export_group(i,Q217_EG1_E2_insp_req_no)
    End If
Next  

Set PQIG040 = Nothing
%>
<Script Language="vbscript">   
With parent
	.ggoSpread.Source = .vspdData 
	.ggoSpread.SSShowDataByClip "<%=strData%>"
	.vspdData.focus
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> "" Then
		 .DbQuery
	Else
		 <% ' Request값을 hidden Varialbe로 넘겨줌 %>
		 .hItemCd	= "<%=ConvSPChars(strItemCd)%>"
		 .hInspReqNo = "<%=ConvSPChars(strInspReqNo)%>"
		 .hBpCd = "<%=ConvSPChars(strBpCd)%>"
		 .hCustCd = "<%=ConvSPChars(strCustCd)%>"
		 .hWcCd = "<%=ConvSPChars(strWcCd)%>"
		 .hFrInspReqDt = "<%=strFrInspReqDt%>"
		 .hToInspReqDt = "<%=strToInspReqDt%>"
		 .DbQueryOk
    End If		
End with
</Script>
