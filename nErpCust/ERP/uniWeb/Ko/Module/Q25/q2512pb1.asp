<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" --> <!--uniDateClientFormate 있을때 만 쓰고 없으면 뺀다.-->
<!-- #Include file="../../inc/IncSvrNumber.inc" --> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2512PB1
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

Dim PQIG270

Dim strData
Dim LngMaxRow
Dim StrNextKey
Dim i

Const C_SHEETMAXROWS_D = 100

'IMPORTS VIEW 상수 
'공통 
Private Const LG_I1_common_plant_cd = 0
Private Const LG_I1_common_insp_req_no = 1
Private Const LG_I1_common_insp_class_cd = 2
Private Const LG_I1_common_item_cd = 3
private Const LG_I1_common_lot_no = 4
Private Const LG_I1_common_insp_status_cd = 5
Private Const LG_I1_common_fr_insp_req_dt = 6
Private Const LG_I1_common_to_insp_req_dt = 7
    
'수입검사 
Private Const LG_I1_r_bp_cd = 8
Private Const LG_I1_r_mvmt_no = 9
Private Const LG_I1_r_por_no = 10
    
'공정검사 
Private Const LG_I1_p_rout_no = 11
Private Const LG_I1_p_opr_no = 12
Private Const LG_I1_p_prodt_no = 13
    
'최종검사 
Private Const LG_I1_f_prodt_no = 14
Private Const LG_I1_f_sl_cd = 15

    
'출하검사 
Private Const LG_I1_s_bp_cd = 16
Private Const LG_I1_s_dn_no = 17

'EXPORTS VIEW 상수 
'공통 
Private Const LG_E1_common_insp_req_no = 0
Private Const LG_E1_common_insp_class_cd = 1
Private Const LG_E1_common_insp_class_nm = 2
Private Const LG_E1_common_item_cd = 3
Private Const LG_E1_common_item_nm = 4
Private Const LG_E1_common_spec = 5
Private Const LG_E1_common_tracking_no = 6
Private Const LG_E1_common_lot_no = 7
Private Const LG_E1_common_lot_sub_no = 8
Private Const LG_E1_common_lot_size = 9
Private Const LG_E1_common_unit_cd = 10
Private Const LG_E1_common_insp_req_dt = 11
Private Const LG_E1_common_insp_reqmt_dt = 12
Private Const LG_E1_common_insp_schdl_dt = 13
Private Const LG_E1_common_insp_status_cd = 14
Private Const LG_E1_common_insp_status_nm = 15
Private Const LG_E1_common_accum_lot_size = 16
Private Const LG_E1_common_if_yesno = 17
    
'수입검사 
Private Const LG_E1_r_bp_cd = 18
Private Const LG_E1_r_bp_nm = 19
Private Const LG_E1_r_mvmt_no = 20
Private Const LG_E1_r_por_no = 21
Private Const LG_E1_r_por_seq = 22
Private Const LG_E1_r_sl_cd = 23
Private Const LG_E1_r_sl_nm = 24    

'공정검사 
Private Const LG_E1_p_rout_no = 25
Private Const LG_E1_p_rout_no_desc = 26
Private Const LG_E1_p_opr_no = 27
Private Const LG_E1_p_opr_no_desc = 28
Private Const LG_E1_p_wc_cd = 29
Private Const LG_E1_p_wc_nm = 30
Private Const LG_E1_p_prodt_no = 31
Private Const LG_E1_p_report_seq = 32
    
'최종검사 
Private Const LG_E1_f_prodt_no = 33
Private Const LG_E1_f_rout_no = 34
Private Const LG_E1_f_rout_no_desc = 35
Private Const LG_E1_f_opr_no = 36
Private Const LG_E1_f_opr_no_desc = 37
Private Const LG_E1_f_report_seq = 38
Private Const LG_E1_f_sl_cd = 39
Private Const LG_E1_f_sl_nm = 40
Private Const LG_E1_f_document_no = 41
Private Const LG_E1_f_document_seq_no = 42
Private Const LG_E1_f_document_sub_no = 43
    
'출하검사 
Private Const LG_E1_s_bp_cd = 44
Private Const LG_E1_s_bp_nm = 45
Private Const LG_E1_s_dn_no = 46
Private Const LG_E1_s_dn_seq = 47

Dim LG_I1_inspection_request

Dim strQueryFlag
Dim LE1_plant_nm 
Dim LE1_item_nm 
Dim LE1_supplier_nm 
Dim LE1_rout_no_desc
Dim LE1_opr_no_desc
Dim LE1_sl_nm 
Dim LE1_bp_nm 

Dim LG_E1_inspection_request

ReDim LG_I1_inspection_request(17)

' MA로 부터 Request Parameter 받기 
LG_I1_inspection_request(LG_I1_common_plant_cd) = Request("txtPlantCd")
LG_I1_inspection_request(LG_I1_common_insp_req_no) = Request("txtInspReqNo")
LG_I1_inspection_request(LG_I1_common_insp_class_cd) = Request("txtInspClassCd")
LG_I1_inspection_request(LG_I1_common_item_cd) = Request("txtItemCd")
LG_I1_inspection_request(LG_I1_common_lot_no) = Request("txtLotNo")
LG_I1_inspection_request(LG_I1_common_insp_status_cd) = Request("txtInspStatusCd")
If Request("txtFrInspReqDt") <> "" Then
	LG_I1_inspection_request(LG_I1_common_fr_insp_req_dt) = UNIConvDate(Request("txtFrInspReqDt"))
End If
If Request("txtToInspReqDt") <> "" Then
	LG_I1_inspection_request(LG_I1_common_to_insp_req_dt) = UNIConvDate(Request("txtToInspReqDt"))
End If

Select Case Request("txtInspClassCd")
	Case "R"
		LG_I1_inspection_request(LG_I1_r_bp_cd) = Request("txtSupplierCd")
		LG_I1_inspection_request(LG_I1_r_mvmt_no) = Request("txtPRNo")
		LG_I1_inspection_request(LG_I1_r_por_no) = Request("txtPONo")	
	Case "P"
		LG_I1_inspection_request(LG_I1_p_rout_no) = Request("txtRoutNo")
		LG_I1_inspection_request(LG_I1_p_opr_no) = Request("txtOprNo")
		LG_I1_inspection_request(LG_I1_p_prodt_no) = Request("txtProdtNo")	
	Case "F"
		LG_I1_inspection_request(LG_I1_f_prodt_no) = Request("txtProdtNo")	
		LG_I1_inspection_request(LG_I1_f_sl_cd) = Request("txtSLCd")
	Case "S"
		LG_I1_inspection_request(LG_I1_s_bp_cd) = Request("txtBpCd")
		LG_I1_inspection_request(LG_I1_s_dn_no) = Request("txtDNNo")	
	Case Else
	
End Select

If Request("lgQueryFlag") = "0" Then		'추가 조회 
	LngMaxRow = CInt(Request("txtMaxRows"))
	StrNextKey = Request("txtInspReqNo")
	strQueryFlag = "A"
Else										'신규 조회 
	LngMaxRow = 0
	StrNextKey = ""
	strQueryFlag = "N"
End If

' 해당 Business Object 생성 
Set PQIG270 = Server.CreateObject("PQIG270.cQListInspRequestSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If


Call PQIG270.Q_LIST_INSP_REQUEST_SVR(gStrGlobalCollection, _
									 C_SHEETMAXROWS_D, _
									 strQueryFlag, _
									 LG_I1_inspection_request, _
									 LE1_plant_nm, _
									 LE1_item_nm, _
									 LE1_supplier_nm, _
									 LE1_rout_no_desc, _
									 LE1_opr_no_desc, _
									 LE1_sl_nm, _
									 LE1_bp_nm, _
									 LG_E1_inspection_request)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG270 = Nothing
	Response.End
End If

For i = 0 To UBound(LG_E1_inspection_request, 1)
    If i < C_SHEETMAXROWS_D Then
    	strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_req_no))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_item_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_item_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_spec))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_bp_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_bp_nm)))
		
		If Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "P" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_rout_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_rout_no_desc))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_opr_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_opr_no_desc))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_wc_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_wc_nm)))
		ElseIf Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "F" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_rout_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_rout_no_desc))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_opr_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_opr_no_desc))) _
							  & Chr(11) & "" _
							  & Chr(11) & ""
		Else
			strData = strData & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & ""
		End If
		
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_s_bp_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_s_bp_nm))) _
			  			  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_tracking_no))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_lot_no)))
						  
		If Trim(LG_E1_inspection_request(i, LG_E1_common_lot_no)) = "" Then
		  	strData = strData & Chr(11) & ""
		Else
		  	strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_common_lot_sub_no), 0, 0)
		End If
						  
		strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_common_lot_size), ggQty.DecPoint, 0) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_unit_cd))) _
						  & Chr(11) & UNIDateClientFormat(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_req_dt))) _
						  & Chr(11) & UNIDateClientFormat(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_reqmt_dt))) _
						  & Chr(11) & UNIDateClientFormat(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_schdl_dt))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_status_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_mvmt_no))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_por_no)))
		
		If Trim(LG_E1_inspection_request(i, LG_E1_r_por_no)) = "" Or IsNull(LG_E1_inspection_request(i, LG_E1_r_por_no)) Then
		  	strData = strData & Chr(11) & ""
		Else
			strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_r_por_seq), 0, 0)
		End If
						  
		If Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "P" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_p_prodt_no)))
			
			If Trim(LG_E1_inspection_request(i, LG_E1_p_prodt_no)) = "" Then
			  	strData = strData & Chr(11) & ""
			Else
				strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_p_report_seq), 0, 0)			  
			End If
			
		ElseIf Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "F" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_prodt_no)))
							  
			If Trim(LG_E1_inspection_request(i, LG_E1_f_prodt_no)) = "" Then
			  	strData = strData & Chr(11) & ""
			Else
				strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_f_report_seq), 0, 0)	  
			End If
		Else
			strData = strData & Chr(11) & "" _
							  & Chr(11) & ""
		End If
		
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_document_no)))
		
		If Trim(LG_E1_inspection_request(i, LG_E1_f_document_no)) = "" Or IsNull(LG_E1_inspection_request(i, LG_E1_f_document_no)) Then
		  	strData = strData & Chr(11) & "" _
		  					  & Chr(11) & ""
		Else
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_document_seq_no))) _		
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_document_sub_no)))	  
		End If
						  
		If Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "R" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_sl_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_r_sl_nm)))
		ElseIf Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd)) = "F" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_sl_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_f_sl_nm)))
		Else
			strData = strData & Chr(11) & "" _
							  & Chr(11) & "" 
		End If
		
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_s_dn_no)))
		If Trim(LG_E1_inspection_request(i, LG_E1_s_dn_no)) = "" Or IsNull(LG_E1_inspection_request(i, LG_E1_s_dn_no)) Then
		  	strData = strData & Chr(11) & ""
		Else
			strData = strData & Chr(11) & UniNumClientFormat(LG_E1_inspection_request(i, LG_E1_s_dn_seq), 0, 0)  
		End If
	    strData = strData & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_class_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E1_inspection_request(i, LG_E1_common_insp_status_cd)))
		strData = strData & Chr(11) & LngMaxRow + i + 1 _
						  & Chr(11) & Chr(12)
						  
    Else
		StrNextKey = ConvSPChars(Trim(LG_E1_inspection_request(i,LG_E1_common_insp_req_no)))
    End If
Next  

Set PQIG270 = Nothing
%>
<Script Language="vbscript">
	With parent
		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		.vspdData.focus
		
		.lgStrPrevKey = "<%=StrNextKey%>"
		
		' 조건부에 명을 보여줌 
		.txtPlantNm.value = "<%=ConvSPChars(Trim(LE1_plant_nm))%>"
		.txtItemNm.value = "<%=ConvSPChars(Trim(LE1_item_nm))%>"
		.txtSupplierNm.value = "<%=ConvSPChars(Trim(LE1_supplier_nm))%>"
		.txtRoutNoDesc.value = "<%=ConvSPChars(Trim(LE1_rout_no_desc))%>"
		.txtOprNoDesc.value = "<%=ConvSPChars(Trim(LE1_opr_no_desc))%>"
		.txtSLNm.value = "<%=ConvSPChars(Trim(LE1_sl_nm))%>"
		.txtBPNm.value = "<%=ConvSPChars(Trim(LE1_bp_nm))%>"
		
		' Request값을 hidden Varialbe로 넘겨줌 
		.hPlantCd = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.hInspReqNo = "<%=ConvSPChars(Request("txtInspReqNo"))%>"
		.hInspClassCd = "<%=ConvSPChars(Request("txtInspClassCd"))%>"
		.hInspStatusCd = "<%=ConvSPChars(Request("txtInspStatusCd"))%>"
		.hItemCd = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.hLotNo = "<%=ConvSPChars(Request("txtLotNo"))%>"
		.hFrInspReqDt = "<%=Request("txtFrInspReqDt")%>"
		.hToInspReqDt = "<%=Request("txtToInspReqDt")%>"
		
		Select Case "<%=Request("txtInspClassCd")%>"
			Case "R"
				.hSupplierCd = "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.hPRNo = "<%=ConvSPChars(Request("txtPRNo"))%>"
				.hPONo = "<%=ConvSPChars(Request("txtPONo"))%>"	
			Case "P"
				.hRoutNo = "<%=ConvSPChars(Request("txtRoutNo"))%>"
				.hOprNo = "<%=ConvSPChars(Request("txtOprNo"))%>"
				.hProdtNo1 = "<%=ConvSPChars(Request("txtProdtNo"))%>"	
			Case "F"
				.hSLCd = "<%=ConvSPChars(Request("txtSLCd"))%>"
				.hProdtNo2 = "<%=ConvSPChars(Request("txtProdtNo"))%>"	
			Case "S"
				.hBPCd = "<%=ConvSPChars(Request("txtBpCd"))%>"
				.hDNNo = "<%=ConvSPChars(Request("txtDNNo"))%>"
			Case Else
			
		End Select
		
		.DbQueryOk
	    
	End with
</Script>
