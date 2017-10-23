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
'*  3. Program ID           : Q2512MB2
'*  4. Program Name         : 검사의뢰 신규/수정 저장 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/05/07
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
On Error Resume Next

Call HideStatusWnd														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

'공통 
Const I1_common_insp_req_no = 0
Const I1_common_plant_cd = 1
Const I1_common_item_cd = 2
Const I1_common_insp_class_cd = 3
Const I1_common_tracking_no = 4
Const I1_common_lot_no = 5
Const I1_common_lot_sub_no = 6
Const I1_common_lot_size = 7
Const I1_common_unit_cd = 8
Const I1_common_insp_req_dt = 9
Const I1_common_insp_reqmt_dt = 10
Const I1_common_insp_schdl_dt = 11

'수입검사 
Const I1_r_bp_cd = 12
Const I1_r_mvmt_no = 13
Const I1_r_por_no = 14
Const I1_r_por_seq = 15
    
'공정검사 
Const I1_p_wc_cd = 16
Const I1_p_opr_no = 17
Const I1_p_prodt_no = 18
Const I1_p_report_seq = 19
    
'최종검사 
Const I1_f_wc_cd = 20
Const I1_f_sl_cd = 21
Const I1_f_document_no = 22
Const I1_f_document_seq_no = 23
Const I1_f_document_sub_no = 24
Const I1_f_prodt_no = 25
Const I1_f_opr_no = 26
Const I1_f_report_seq = 27
    
'출하검사 
Const I1_s_bp_cd = 28
Const I1_s_dn_no = 29
Const I1_s_dn_seq = 30

'수입검사 
Const I1_r_sl_cd = 31
Const I1_r_sl_nm = 32
Const I1_r_bp_nm = 33

'공정검사 
Const I1_p_rout_no = 34
Const I1_p_rout_no_desc = 35
Const I1_p_wc_nm = 36
Const I1_p_opr_no_desc = 37
    
'최종검사 
Const I1_f_rout_no = 38
Const I1_f_rout_no_desc = 39
Const I1_f_wc_nm = 40
Const I1_f_sl_nm = 41
Const I1_f_opr_no_desc = 42
    
'출하검사 
Const I1_s_bp_nm = 43
    
Dim objPQIG250																	
Dim lgIntFlgMode	
	
Dim sCommandSent
Dim I1_q_inspection_request
Dim E1_insp_req_no
	
lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
If lgIntFlgMode = OPMD_CMODE Then
	sCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	sCommandSent = "UPDATE"
End If
    
ReDim I1_q_inspection_request(43)
    
I1_q_inspection_request(I1_common_insp_req_no) = UCase(Trim(Request("txtInspReqNo2")))
I1_q_inspection_request(I1_common_plant_cd) = UCase(Trim(Request("txtPlantCd")))
I1_q_inspection_request(I1_common_item_cd) = UCase(Trim(Request("txtItemCd")))
I1_q_inspection_request(I1_common_insp_class_cd) = UCase(Request("cboInspClassCd"))
I1_q_inspection_request(I1_common_tracking_no) = UCase(Trim(Request("txtTrackingNo")))
I1_q_inspection_request(I1_common_lot_no) = UCase(Trim(Request("txtLotNo")))
If Request("txtLotSubNo") <> "" Then
	I1_q_inspection_request(I1_common_lot_sub_no) = UNIConvNum(Request("txtLotSubNo"), 0)
End If
I1_q_inspection_request(I1_common_lot_size) = UNIConvNum(Request("txtLotSize"), 0)
I1_q_inspection_request(I1_common_unit_cd) = UCase(Trim(Request("txtUnit")))
I1_q_inspection_request(I1_common_insp_req_dt) = UniConvDate(Request("txtInspReqDt"))
	
If UniConvDate(Request("txtInspReqmtDt")) <> "" Then
	I1_q_inspection_request(I1_common_insp_reqmt_dt) = UniConvDate(Request("txtInspReqmtDt"))
End If
If UniConvDate(Request("txtInspSchdlDt")) <> "" Then
	I1_q_inspection_request(I1_common_insp_schdl_dt) = UniConvDate(Request("txtInspSchdlDt"))
End If
	
Select Case Request("cboInspClassCd")
	Case "R"
		I1_q_inspection_request(I1_r_bp_cd) = UCase(Trim(Request("txtSupplierCd")))
		I1_q_inspection_request(I1_r_bp_nm) = Trim(Request("txtSupplierNm"))
		I1_q_inspection_request(I1_r_mvmt_no) = UCase(Trim(Request("txtPRNo")))
		I1_q_inspection_request(I1_r_por_no) = UCase(Trim(Request("txtPONo")))
		If Request("txtPOSeq") <> "" Then
			I1_q_inspection_request(I1_r_por_seq) = UNIConvNum(Request("txtPOSeq"), 0)
		End If
		I1_q_inspection_request(I1_r_sl_cd) = UCase(Trim(Request("txtSLCd1")))
		I1_q_inspection_request(I1_r_sl_nm) = Trim(Request("txtSLNm1"))
		
	Case "P"
		I1_q_inspection_request(I1_p_wc_cd) = UCase(Trim(Request("txtWcCd")))
		I1_q_inspection_request(I1_p_wc_nm) = Trim(Request("txtWcNm"))
		I1_q_inspection_request(I1_p_rout_no) = UCase(Trim(Request("txtRoutNoforP")))
		I1_q_inspection_request(I1_p_rout_no_desc) = Trim(Request("txtRoutNoDescforP"))
		I1_q_inspection_request(I1_p_opr_no) = UCase(Trim(Request("txtOprNoforP")))
		I1_q_inspection_request(I1_p_opr_no_desc) = UCase(Trim(Request("txtOprNoDescforP")))
		I1_q_inspection_request(I1_p_prodt_no) = UCase(Trim(Request("txtProdtNo1")))
			
	Case "F"
		I1_q_inspection_request(I1_f_sl_cd) = UCase(Trim(Request("txtSLCd2")))
		I1_q_inspection_request(I1_f_sl_nm) = Trim(Request("txtSLNm2"))
		I1_q_inspection_request(I1_f_prodt_no) = UCase(Trim(Request("txtProdtNo2")))
		I1_q_inspection_request(I1_f_rout_no) = UCase(Trim(Request("txtRoutNoforF")))
		I1_q_inspection_request(I1_f_rout_no_desc) = Trim(Request("txtRoutNoDescforF"))
		I1_q_inspection_request(I1_f_opr_no) = UCase(Trim(Request("txtOprNoforF")))
		I1_q_inspection_request(I1_f_opr_no_desc) = UCase(Trim(Request("txtOprNoDescforF")))
	Case "S"
		I1_q_inspection_request(I1_s_bp_cd) = UCase(Trim(Request("txtBPCd")))
		I1_q_inspection_request(I1_s_bp_nm) = Trim(Request("txtBPNm"))
		I1_q_inspection_request(I1_s_dn_no) = UCase(Trim(Request("txtDNNo")))
		If Request("txtDNSeq") <> "" Then
			I1_q_inspection_request(I1_s_dn_seq) = UNIConvNum(Request("txtDNSeq"), 0)
		End If
End Select
	
Set objPQIG250 = Server.CreateObject("PQIG250.cQMaintInspRequestSvr")
	
If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
	
Call objPQIG250.Q_MAINT_INSP_REQUEST_SVR(gStrGlobalCollection, sCommandSent, I1_q_inspection_request, "N", E1_insp_req_no)
If CheckSYSTEMError(Err,True) = true Then
   Set objPQIG250 = Nothing
   Response.End
End if
	
Set objPQIG250 = Nothing
%>
<Script Language=vbscript>
With parent
	If "<%=E1_insp_req_no%>" <> "" Then
		.frm1.txtInspReqNo2.value = "<%=E1_insp_req_no%>"
	End If
	.DbSaveOk
End With
</Script>
