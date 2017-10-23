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
'*  3. Program ID           : Q2511MA1
'*  4. Program Name         : 검사의뢰조회 
'*  5. Program Desc         : 
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf
								
On Error Resume Next
Call HideStatusWnd 

Dim PQIG040													  '☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													  '☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          

Dim strPlantCd
Dim strInspClassCd
Dim strInspReqNo
Dim strItemCd
Dim strBpCd
Dim strCustCd
Dim strWcCd
Dim strInspReqDtFr
Dim strInspReqDtTo
Dim strInspStatus
Dim strData
Dim PvArr
Dim EG1_export_group

Const C_SHEETMAXROWS_D = 100

Dim I5_q_inspection_request
ReDim I5_q_inspection_request(4)
Const Q217_I5_insp_req_no = 0
Const Q217_I5_insp_class_cd = 1
Const Q217_I5_bp_cd = 2
Const Q217_I5_wc_cd = 3
Const Q217_I5_insp_status = 4

Dim E1_b_plant
ReDim E1_b_plant(1)
Const Q217_E1_plant_cd = 0
Const Q217_E1_plant_nm = 1

Const Q217_EG1_E1_plant_cd = 0
Const Q217_EG1_E1_plant_nm = 1
Const Q217_EG1_E2_insp_req_no = 2
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
Const Q217_EG1_E3_item_cd = 19
Const Q217_EG1_E3_item_nm = 20
Const Q217_EG1_E4_wc_nm = 21
Const Q217_EG1_E5_bp_nm = 22
Const Q217_EG1_E6_minor_nm = 23



lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")

strPlantCd = Request("txtPlantCd")
strInspClassCd = Request("cboInspClassCd")
strInspReqNo = Request("txtInspReqNo")
strItemCd = Request("txtItemCd")
strBpCd	= Request("txtBpCd")
strCustCd = Request("txtCustCd")
strWcCd	= Request("txtWcCd")
strInspReqDtFr = Request("txtInspReqDtFr")
strInspReqDtTo = Request("txtInspReqDtTo")
strInspStatus =  Request("cboStatusFlag")


Set PQIG040 = Server.CreateObject("PQIG040.cQListInspReqSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If


If lgStrPrevKey = "" Then
	I5_q_inspection_request(Q217_I5_insp_req_no) = strInspReqNo
Else
	I5_q_inspection_request(Q217_I5_insp_req_no) = lgStrPrevKey
End If

I5_q_inspection_request(Q217_I5_insp_class_cd) = strInspClassCd

Select Case strInspClassCd
	Case "R"
		I5_q_inspection_request(Q217_I5_bp_cd) = strBpCd	
	Case "S"
		I5_q_inspection_request(Q217_I5_bp_cd) = strCustCd
End Select 

I5_q_inspection_request(Q217_I5_wc_cd) = strWcCd

If strInspStatus <> "" Then
	I5_q_inspection_request(Q217_I5_insp_status) = strInspStatus
End If

If strInspReqDtFr <> "" Then 
	strInspReqDtFr = UNIConvDate(strInspReqDtFr)
End If
If strInspReqDtTo <> "" Then 
	strInspReqDtTo = UNIConvDate(strInspReqDtTo)
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

ReDim PvArr(UBound(EG1_export_group, 1))

For LngRow = 0 To UBound(EG1_export_group, 1)
    If LngRow < C_SHEETMAXROWS_D Then
		PvArr(LngRow) = Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_insp_req_no))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E3_item_cd))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E3_item_nm))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_bp_cd))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E5_bp_nm))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_wc_cd))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E4_wc_nm))) & _
						Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_insp_req_dt)))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_lot_no))) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_lot_sub_no))) & _
						Chr(11) & UniNumClientFormat(Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E2_lot_size))), ggQty.DecPoint ,0) & _
						Chr(11) & Trim(ConvSPChars(EG1_export_group(LngRow, Q217_EG1_E6_minor_nm))) & _
						Chr(11) & LngMaxRow + LngRow + 1 & Chr(11) & Chr(12)
    Else
		StrNextKey = EG1_export_group(LngRow, Q217_EG1_E2_insp_req_no)
    End If
Next
strData = Join(PvArr, "")

Set PQIG040 = Nothing
%>
<Script Language=vbscript>
With Parent
	'Header
	.frm1.txtPlantCd.Value = "<%=ConvSPChars(E1_b_plant(Q217_E1_plant_cd))%>"
	.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_b_plant(Q217_E1_plant_nm))%>"
    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=strData%>"
		
	.lgStrPrevKey = "<%=StrNextKey%>"
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else			
		 <% ' Request값을 hidden input으로 넘겨줌 %>
		 .frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		 .frm1.hInspClassCd.value = "<%=ConvSPChars(strInspClassCd)%>"
		 .frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
		 .frm1.hItemCd.value = "<%=ConvSPChars(strItemCd)%>"
		 .frm1.hBpCd.value = "<%=ConvSPChars(strBpCd)%>"
		 .frm1.hCustCd.value = "<%=ConvSPChars(strCustCd)%>"
		 .frm1.hWcCd.value = "<%=ConvSPChars(strWcCd)%>"			 
		 .frm1.hInspReqDtFr.Value = "<%=strInspReqDtFr%>"	 
		 .frm1.hInspReqDtTo.Value = "<%=strInspReqDtTo%>"
		 .frm1.hStatusFlag.Value = "<%=ConvSPChars(strInspStatus)%>"
		 .DbQueryOk
    End If		
End with
</Script>	
<%
Set PQIG040 = Nothing
%>
