<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4112MB1
'*  4. Program Name         : 부적합처리조회 
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

On Error Resume Next		

Call HideStatusWnd

Const C_SHEETMAXROWS_D = 100

Dim PQIG330													'☆ : 조회용 ComProxy Dll 사용 변수 

Dim StrNextKey		' 다음 값 

Dim lgAddQueryFlag
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow

Dim StrPlantCd
Dim StrInspReqNo
Dim IntInspResultNo
Dim StrDispositionCd

Dim StrData

'EXPORTS VIEW
'PLANT
Const Q330_E1_plant_cd = 0
Const Q330_E1_plant_nm = 1

'INSPECTION RESULT
Const Q330_E2_insp_req_no = 0
Const Q330_E2_insp_class_cd = 1
Const Q330_E2_insp_class_nm = 2
Const Q330_E2_item_cd = 3
Const Q330_E2_item_nm = 4
Const Q330_E2_spec = 5
Const Q330_E2_lot_no = 6
Const Q330_E2_lot_sub_no = 7
Const Q330_E2_lot_size = 8
Const Q330_E2_Unit_cd = 9
Const Q330_E2_insp_qty = 10
Const Q330_E2_defect_qty = 11
Const Q330_E2_decision = 12
Const Q330_E2_decision_nm = 13
Const Q330_E2_status_flag = 14

'수입검사 
Const Q330_E2_r_bp_cd = 15
Const Q330_E2_r_bp_nm = 16
Const Q330_E2_r_sl_cd = 17
Const Q330_E2_r_sl_nm = 18
'공정검사 
Const Q330_E2_p_rout_no = 19
Const Q330_E2_p_rout_no_desc = 20
Const Q330_E2_p_opr_no = 21
Const Q330_E2_p_opr_no_desc = 22
Const Q330_E2_p_wc_cd = 23
Const Q330_E2_p_wc_nm = 24

'최종검사 
Const Q330_E2_f_sl_cd = 25
Const Q330_E2_f_sl_nm = 26

'출하검사 
Const Q330_E2_s_bp_cd = 27
Const Q330_E2_s_bp_nm = 28

'INSPECTION DISPOSITION
Const Q330_E3_disposition_cd = 0
Const Q330_E3_disposition_nm = 1
Const Q330_E3_qty = 2
Const Q330_E3_remark = 3

Dim iExportPlant
Dim iExportInspectionResult
Dim iExportInspectionDisposition
Dim iExportDispositionNFError

StrPlantCd = Request("txtPlantCd")
StrInspReqNo = Request("txtInspReqNo")
IntInspResultNo = 1
lgAddQueryFlag 	= Request("lgAddQueryFlag")

If lgAddQueryFlag = False Then
	StrDispositionCd = ""
	LngMaxRow = 0	
Else
	StrDispositionCd = Request("txtDispositionCd")
	LngMaxRow = CLng(Request("txtMaxRows"))
End If

Set PQIG330 = Server.CreateObject("PQIG330.cQLiInspDispSimple")
If CheckSystemError(Err,True) Then						
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PQIG330.Q_LIST_INSP_DISPOSIT_SIMPLE_SVR(gStrGlobalCollection, _
										C_SHEETMAXROWS_D, _
										StrPlantCd, _
										StrInspReqNo, _
										IntInspResultNo , _
										StrDispositionCd , _
										lgAddQueryFlag , _
										iExportPlant , _
										iExportInspectionResult , _
										iExportInspectionDisposition, _
										iExportDispositionNFError )

If CheckSystemError(Err,True) Then
	Set PQIG330= Nothing
	Response.End
End If

Set PQIG330= Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	'Header
	.txtPlantCd.value = "<%=ConvSPChars(Trim(iExportPlant(Q330_E1_plant_cd)))%>"
	.txtPlantNm.value = "<%=ConvSPChars(Trim(iExportPlant(Q330_E1_plant_nm)))%>"
	.txtInspReqNo.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_insp_req_no)))%>"
	If <%=lgAddQueryFlag%> = False Then
		.txtInspClassNm.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_insp_class_nm)))%>"
		.txtDecision.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_decision_nm)))%>"
		.txtItemCd.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_item_cd)))%>"
		.txtItemNm.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_item_nm)))%>"
		.txtSpec.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_spec)))%>"		
		.txtLotNo.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_lot_no)))%>"
		If "<%=ConvSPChars(Trim(iExportInspectionResult(Q330_E2_lot_no)))%>" <> "" Then
			.txtLotSubNo.value = "<%=UniNumClientFormat(iExportInspectionResult(Q330_E2_lot_sub_no), 0, 0)%>"
		End If
		
		.txtLotSize.Text = "<%=UniNumClientFormat(iExportInspectionResult(Q330_E2_lot_size), ggQty.DecPoint, 0)%>"
		.txtUnit.value = "<%=iExportInspectionResult(Q330_E2_Unit_cd)%>"
		.txtInspQty.Text = "<%=UniNumClientFormat(iExportInspectionResult(Q330_E2_insp_qty), ggQty.DecPoint, 0)%>"
		.txtDefectQty.Text = "<%=UniNumClientFormat(iExportInspectionResult(Q330_E2_defect_qty), ggQty.DecPoint, 0)%>"
		.hInspClassCd.value = "<%=UCase(Trim(iExportInspectionResult(Q330_E2_insp_class_cd)))%>"
		.hDecisionCd.value = "<%=UCase(Trim(iExportInspectionResult(Q330_E2_decision)))%>"
		.hStatusFlag.value = "<%=UCase(Trim(iExportInspectionResult(Q330_E2_status_flag)))%>"

		Select Case "<%=UCase(Trim(iExportInspectionResult(Q330_E2_insp_class_cd)))%>"
			Case "R"
				.txtSupplierCd.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_r_bp_cd))%>"
				.txtSupplierNm.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_r_bp_nm))%>"
				.txtSLCd1.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_r_sl_cd))%>"
				.txtSLNm1.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_r_sl_nm))%>"	
			Case "P"
				.txtRoutNo.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_rout_no))%>"
				.txtRoutNoDesc.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_rout_no_desc))%>"
				.txtOprNo.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_opr_no))%>"
				.txtOprNoDesc.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_opr_no_desc))%>"
				.txtWcCd.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_wc_cd))%>"
				.txtWcNm.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_p_wc_nm))%>"

			Case "F"
				.txtSLCd2.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_f_sl_cd))%>"
				.txtSLNm2.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_f_sl_nm))%>"

			Case "S"
				.txtBPCd.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_s_bp_cd))%>"
				.txtBPNm.value = "<%=ConvSPChars(iExportInspectionResult(Q330_E2_s_bp_nm))%>"
		End Select
				
	End If
End With		 
</SCRIPT>
<%    
If iExportDispositionNFError <> "" Then
	Call DisplayMsgBox(iExportDispositionNFError, vbOKOnly, "", "", I_MKSCRIPT)
%>
<Script Language=vbscript>
	With Parent
	'	Call .SetToolBar("11100000000111")
		Call .ChangingFieldByInspClass(.frm1.hInspClassCd.value)
	End With
</Script>
<%
	Response.End 
End If
 
For LngRow = 0 To UBound(iExportInspectionDisposition)
	If LngRow < C_SHEETMAXROWS_D Then 
		StrData = StrData & Chr(11) & ConvSPChars(Trim(iExportInspectionDisposition(LngRow, Q330_E3_disposition_cd)))
		StrData = StrData & Chr(11) & ""
		StrData = StrData & Chr(11) & ConvSPChars(Trim(iExportInspectionDisposition(LngRow, Q330_E3_disposition_nm)))
		StrData = StrData & Chr(11) & UniNumClientFormat(iExportInspectionDisposition(LngRow, Q330_E3_qty), ggQty.DecPoint ,0)				'qty
		StrData = StrData & Chr(11) & ConvSPChars(Trim(iExportInspectionDisposition(LngRow, Q330_E3_remark)))	'remark
		StrData = StrData & Chr(11) & LngMaxRow + LngRow + 1
		StrData = StrData & Chr(11) & Chr(12)
	Else
		StrNextKey = ConvSPChars(Trim(iExportInspectionDisposition(LngRow, Q330_E3_disposition_cd)))
	End if
Next
%>
<Script Language=vbscript>
	With Parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip "<%=StrData%>"
		
		.lgStrPrevKey = "<%=StrNextKey%>"
		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
			.DbQuery
		Else
			<% ' Request값을 hidden input으로 넘겨줌 %>
			.frm1.hInspReqNo.value = "<%=ConvSPChars(StrInspReqNo)%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(StrPlantCd)%>"
			.DbQueryOk
		End if
	End with
</Script>	
