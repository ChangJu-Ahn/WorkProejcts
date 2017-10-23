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
'*  3. Program ID           : Q2311MB5
'*  4. Program Name         : 검사등록 
'*  5. Program Desc         : 검사의뢰번호 LOOK UP
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
	
'EXPORTS VIEW 상수 
'B_Plant
Const E1_plant_cd = 0
Const E1_plant_nm = 1
    
'Q_Inspection_Request
'공통 
Const E2_common_insp_req_no = 0
Const E2_common_item_cd = 1
Const E2_common_item_nm = 2
Const E2_common_item_spec = 3
Const E2_common_insp_class_cd = 4
Const E2_common_insp_class_nm = 5
Const E2_common_tracking_no = 6
Const E2_common_lot_no = 7
Const E2_common_lot_sub_no = 8
Const E2_common_lot_size = 9
Const E2_common_unit_cd = 10
Const E2_common_insp_req_dt = 11
Const E2_common_insp_reqmt_dt = 12
Const E2_common_insp_schdl_dt = 13
Const E2_common_insp_status_cd = 14
Const E2_common_insp_status_nm = 15
Const E2_common_accum_lot_size = 16
Const E2_common_if_yesno = 17

'수입검사 
Const E2_r_bp_cd = 18
Const E2_r_bp_nm = 19
Const E2_r_mvmt_no = 20
Const E2_r_por_no = 21
Const E2_r_por_seq = 22
Const E2_r_sl_cd = 23
Const E2_r_sl_nm = 24
    
'공정검사 
Const E2_p_rout_no = 25
Const E2_p_rout_no_desc = 26
Const E2_p_opr_no = 27
Const E2_p_opr_no_desc = 28
Const E2_p_wc_cd = 29
Const E2_p_wc_nm = 30
Const E2_p_prodt_no = 31
Const E2_p_report_seq = 32
    
'최종검사 
Const E2_f_prodt_no = 33
Const E2_f_rout_no = 34
Const E2_f_rout_no_desc = 35
Const E2_f_opr_no = 36
Const E2_f_opr_no_desc = 37
Const E2_f_report_seq = 38
Const E2_f_sl_cd = 39
Const E2_f_sl_nm = 40
Const E2_f_document_no = 41
Const E2_f_document_seq_no = 42
Const E2_f_document_sub_no = 43
    
'출하검사 
Const E2_s_bp_cd = 44
Const E2_s_bp_nm = 45
Const E2_s_dn_no = 46
Const E2_s_dn_seq = 47
	
Dim objPQIG260
Dim strPlantCd
Dim strInspReqNo
Dim strPrevNextFlg
Dim EG1_b_plant
Dim EG2_q_inspection_request
Dim E3_PrevNextError
	
Set objPQIG260 = Server.CreateObject("PQIG260.cQLookUpInspRequestSvr")    

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
	    
strPlantCd = Request("txtPlantCd")
strInspReqNo = Request("txtInspReqNo")
	
Call objPQIG260.Q_LOOK_UP_INSP_REQUEST_SVR(gStrGlobalCollection, _
											"", _
											strPlantCd, _
											strInspReqNo, _
											"N", _
											EG1_b_plant, _
											EG2_q_inspection_request, _
											E3_PrevNextError)
	
If CheckSYSTEMError(Err,True) = true Then
   Set objPQIG260 = Nothing
   Response.End
End if
Set objPQIG260 = Nothing

If UCase(EG2_q_inspection_request(E2_common_insp_class_cd)) <> "F" Then
%>
<Script Language=vbscript>
	With Parent
		Call .DisplayMsgBox("223701","X","X","X") 		'검사의뢰 정보에 자료가 없습니다 
	End With
</Script>
<%
	Response.End
End If

SELECT CASE Trim(EG2_q_inspection_request(E2_common_insp_status_cd))
	CASE "I"
		Call DisplayMsgBox("223716", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	CASE "D"
		Call DisplayMsgBox("223717", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	CASE "R"
		Call DisplayMsgBox("223718", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	CASE ELSE
END SELECT
%>
<Script Language=vbscript>
With parent.frm1		
	'content
	
	.txtInspReqNo2.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_common_insp_req_no)))%>"
	.hInspReqNo2.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_common_insp_req_no)))%>"
	.txtItemCd.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_common_item_cd)))%>"
	.txtItemNm.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_common_item_nm)))%>"
	
	.txtLotNo.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_common_lot_no)))%>"
	If "<%=EG2_q_inspection_request(E2_common_lot_sub_no)%>" <> ""  Then
		.txtLotSubNo.value = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_common_lot_sub_no), 0 ,0)%>"
	End If
	
	.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_common_lot_size), ggQty.DecPoint ,0)%>"
End With
</Script>
