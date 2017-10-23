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
'*  3. Program ID           : Q4111PB1
'*  4. Program Name         : 검사결과현황 팝업 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/07/01
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next		

Call HideStatusWnd

Dim PQIG300

Dim LngMaxRow
Dim StrNextKey1
Dim StrNextKey2
Dim i
Dim StrData
Dim strQueryFlag

Dim LG_I1_inspection_result
Dim LG_E2_inspection_result

Dim LE1_plant_nm 
Dim LE1_item_nm 
Dim LE1_supplier_nm 
Dim LE1_rout_no_desc
Dim LE1_opr_no_desc
Dim LE1_sl_nm 
Dim LE1_bp_nm 

Const C_SHEETMAXROWS_D = 100

'IMPORTS VIEW 상수 
'공통 
Const I1_plant_cd = 0
Const I1_insp_req_no = 1
Const I1_insp_result_no = 2
Const I1_insp_class_cd = 3
Const I1_item_cd = 4
Const I1_lot_no = 5
Const I1_fr_insp_dt = 6
Const I1_to_insp_dt = 7
Const I1_status_flag_cd = 8
Const I1_decision_cd = 9
'수입검사 
Const I1_r_bp_cd = 10
    
'공정검사 
Const I1_p_rout_no = 11
Const I1_p_opr_no = 12
    
'최종검사 
Const I1_f_sl_cd = 13
    
'출하검사 
Const I1_s_bp_cd = 14

'EXPORTS VIEW 상수 
'공통 
Const E2_insp_req_no = 0
Const E2_insp_result_no = 1
Const E2_insp_class_cd = 2
Const E2_insp_class_nm = 3
Const E2_item_cd = 4
Const E2_item_nm = 5
Const E2_spec = 6
Const E2_tracking_no = 7
Const E2_lot_no = 8
Const E2_lot_sub_no = 9
Const E2_lot_size = 10
Const E2_unit_cd = 11
Const E2_inspector_cd = 12
Const E2_inspector_nm = 13
Const E2_insp_dt = 14
Const E2_insp_qty = 15
Const E2_defect_qty = 16
Const E2_decision_cd = 17
Const E2_decision_nm = 18
Const E2_status_flag_cd = 19
Const E2_status_flag_nm = 20
Const E2_remark = 21
    
'수입검사 
Const E2_r_bp_cd = 22
Const E2_r_bp_nm = 23
    
'공정검사 
Const E2_p_prodt_order_no = 24
Const E2_p_rout_no = 25
Const E2_p_rout_no_desc = 26
Const E2_p_opr_no = 27
Const E2_p_opr_no_desc = 28
Const E2_p_wc_cd = 29
Const E2_p_wc_nm = 30
    
'최종검사 
Const E2_f_sl_cd		= 31
Const E2_f_sl_nm		= 32
Const E2_f_sl_cd_good	= 33
Const E2_f_sl_nm_good	= 34
Const E2_f_sl_cd_defect = 35
Const E2_f_sl_nm_defect = 36
    
'출하검사 
Const E2_s_bp_cd = 37
Const E2_s_bp_nm = 38

'수입검사 
Const E2_r_sl_cd		= 39
Const E2_r_sl_nm		= 40
Const E2_r_sl_cd_good	= 41
Const E2_r_sl_nm_good	= 42
Const E2_r_sl_cd_defect = 43
Const E2_r_sl_nm_defect = 44

Const E2_release_dt = 45

ReDim LG_I1_inspection_result(14)

' MA로 부터 Request Parameter 받기 
LG_I1_inspection_result(I1_plant_cd) = Request("txtPlantCd")
LG_I1_inspection_result(I1_insp_req_no) = Request("txtInspReqNo")
LG_I1_inspection_result(I1_insp_result_no) = Request("txtInspResultNo")
LG_I1_inspection_result(I1_insp_class_cd) = Request("txtInspClassCd")
LG_I1_inspection_result(I1_item_cd) = Request("txtItemCd")
LG_I1_inspection_result(I1_lot_no) = Request("txtLotNo")
LG_I1_inspection_result(I1_status_flag_cd) = Request("txtStatusFlagCd")
LG_I1_inspection_result(I1_decision_cd) = Request("txtDecisionCd")

If Request("txtFrInspDt") <> "" Then
	LG_I1_inspection_result(I1_fr_insp_dt) = UNIConvDate(Request("txtFrInspDt"))
End If
If Request("txtToInspDt") <> "" Then
	LG_I1_inspection_result(I1_to_insp_dt) = UNIConvDate(Request("txtToInspDt"))
End If
Select Case Request("txtInspClassCd")
	Case "R"
		LG_I1_inspection_result(I1_r_bp_cd) = Request("txtSupplierCd")
		
	Case "P"
		LG_I1_inspection_result(I1_p_rout_no) = Request("txtRoutNo")
		LG_I1_inspection_result(I1_p_opr_no) = Request("txtOprNo")
		
	Case "F"
		LG_I1_inspection_result(I1_f_sl_cd) = Request("txtSLCd")
		
	Case "S"
		LG_I1_inspection_result(I1_s_bp_cd) = Request("txtBpCd")
		
	Case Else
	
End Select

If Request("lgQueryFlag") = "0" Then		'추가 조회 
	LngMaxRow = Request("txtMaxRows")
	StrNextKey1 = Request("txtInspReqNo")
	StrNextKey2 = Request("txtInspResultNo")
	strQueryFlag = "A"
Else										'신규 조회 
	LngMaxRow = 0
	StrNextKey1 = ""
	StrNextKey2 = ""
	strQueryFlag = "N"
End If

' 해당 Business Object 생성 
Set PQIG300 = Server.CreateObject("PQIG300.cQLiInspResultSimple")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG300.Q_LIST_INSP_RESULT_SIMPLE_SVR(gStrGlobalCollection, _
									 C_SHEETMAXROWS_D, _
									 strQueryFlag, _
									 LG_I1_inspection_result, _
									 LE1_plant_nm, _
									 LE1_item_nm, _
									 LE1_supplier_nm, _
									 LE1_rout_no_desc, _
									 LE1_opr_no_desc, _
									 LE1_sl_nm, _
									 LE1_bp_nm, _
									 LG_E2_inspection_result)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG300 = Nothing
	Response.End
End If

For i = 0 To UBound(LG_E2_inspection_result, 1)
    If i < C_SHEETMAXROWS_D Then
    	strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_insp_req_no)))

		If Trim(LG_E2_inspection_result(i, E2_insp_class_cd)) = "P" Or Trim(LG_E2_inspection_result(i, E2_insp_class_cd)) = "F" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_prodt_order_no)))
		Else
			strData = strData & Chr(11) & ""
		End If
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_tracking_no))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_insp_class_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_item_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_item_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_spec))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_bp_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_bp_nm)))
		
		If Trim(LG_E2_inspection_result(i, E2_insp_class_cd)) = "P" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_rout_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_rout_no_desc))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_opr_no))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_opr_no_desc))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_wc_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_p_wc_nm)))
		Else
			strData = strData & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & ""
		End If
		
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_s_bp_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_s_bp_nm)))
						  
		If Trim(LG_E2_inspection_result(i, E2_insp_class_cd)) = "R" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_nm))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_cd_good))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_nm_good))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_cd_defect))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_r_sl_nm_defect)))
		ElseIf Trim(LG_E2_inspection_result(i, LG_E1_insp_class_cd)) = "F" Then
			strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_cd))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_nm))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_cd_good))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_nm_good))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_cd_defect))) _
							  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_f_sl_nm_defect)))
		Else
			strData = strData & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & "" _
							  & Chr(11) & ""
		End If			
							  
		strData = strData & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_status_flag_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_decision_nm))) _
						  & Chr(11) & UNIDateClientFormat(Trim(LG_E2_inspection_result(i, E2_insp_dt))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_lot_no)))
		
		If Trim(LG_E2_inspection_result(i, E2_lot_no)) = "" Or IsNull(LG_E2_inspection_result(i, E2_lot_no)) Then
			strData = strData & Chr(11) & ""
		Else
			strData = strData & Chr(11) & UniNumClientFormat(LG_E2_inspection_result(i, E2_lot_sub_no), 0, 0)
		End If
		
		strData = strData & Chr(11) & UniNumClientFormat(LG_E2_inspection_result(i, E2_lot_size), ggQty.DecPoint, 0) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_unit_cd))) _
						  & Chr(11) & UniNumClientFormat(LG_E2_inspection_result(i, E2_insp_qty), ggQty.DecPoint, 0) _
						  & Chr(11) & UniNumClientFormat(LG_E2_inspection_result(i, E2_defect_qty), ggQty.DecPoint, 0) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_inspector_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_inspector_nm))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_insp_class_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_status_flag_cd))) _
						  & Chr(11) & ConvSPChars(Trim(LG_E2_inspection_result(i, E2_decision_cd))) _
						  & Chr(11) & UNIDateClientFormat(Trim(LG_E2_inspection_result(i, E2_release_dt))) _
						  & Chr(11) & LngMaxRow + i + 1 _
						  & Chr(11) & Chr(12)
						  
    Else
		StrNextKey1 = ConvSPChars(Trim(LG_E2_inspection_result(i, E2_insp_req_no)))
		StrNextKey2 = ConvSPChars(Trim(LG_E2_inspection_result(i, E2_insp_result_no)))
    End If
Next  

Set PQIG300 = Nothing
%>
<Script Language="vbscript">   
    With parent
		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		.vspdData.focus
		
		.lgStrPrevKey1 = "<%=StrNextKey1%>"
		.lgStrPrevKey2 = "<%=StrNextKey2%>"
		
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
		
		.hItemCd = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.hLotNo = "<%=ConvSPChars(Request("txtLotNo"))%>"
		.hFrInspDt = "<%=Request("txtFrInspDt")%>"
		.hToInspDt = "<%=Request("txtToInspDt")%>"
		.hStatusFlagCd = "<%=ConvSPChars(Request("txtStatusFlagCd"))%>"
		.hDecisionCd = "<%=ConvSPChars(Request("txtDecisionCd"))%>"
		
		Select Case "<%=Request("txtInspClassCd")%>"
			Case "R"
				.hSupplierCd = "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				
			Case "P"
				.hRoutNo = "<%=ConvSPChars(Request("txtRoutNo"))%>"
				.hOprNo = "<%=ConvSPChars(Request("txtOprNo"))%>"
				
			Case "F"
				.hSLCd = "<%=ConvSPChars(Request("txtSLCd"))%>"
				
			Case "S"
				.hBPCd = "<%=ConvSPChars(Request("txtBpCd"))%>"
				
			Case Else
			
		End Select
		
		.DbQueryOk
	    
	End with
</Script>
