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
'*  3. Program ID           : Q2311MB4
'*  4. Program Name         : 검사등록 
'*  5. Program Desc         : 품목별 검사기준 조회 
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
													
Dim LngRow
Dim LngMaxRow
Dim PQIG030

Dim strInspReqNo
Dim strLotSize
Dim I2_q_inspection_result
Dim EG1_export_group
Dim E1,E2

Const EG1_Insp_Item_Cd =0
Const EG1_Insp_Item_Nm = 1
Const EG1_Insp_Method_Cd = 2
Const EG1_Weight_Cd = 3
Const EG1_insp_spec=4
Const EG1_usl=5
Const EG1_lsl=6
Const EG1_measmt_unit_cd =7
Const EG1_ucl =8
Const EG1_lcl =9
Const EG1_insp_order =10
Const EG1_insp_unit_indctn = 11
Const EG1_InspMethodCd_nm = 12
Const EG1_measmt_equipmt_cd = 13
Const EG1_measmt_equipmt_nm = 14
Const EG1_insp_series = 15
Const EG1_sample_qty = 16
Const EG1_accpt_decision_qty = 17
Const EG1_rejt_decision_qty = 18
Const EG1_accpt_decision_discreate = 19
Const EG1_max_defect_ratio = 20
Const EG1_InspUnitIndctn_nm = 21

Redim I2_q_inspection_result(1)

strInspReqNo = Request("txtInspReqNo")
strLotSize = Request("txtLotSize")

Set PQIG030 = Server.CreateObject("PQIG030.cQListInspItemInsp")
If CheckSystemError(Err,True) Then
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

I2_q_inspection_result(0) =1
I2_q_inspection_result(1) =  UniConvNum(strLotSize, 0)

Call PQIG030.Q_LIST_INSP_ITEM_FOR_INSP(gstrGlobalCollection,strInspReqNo,I2_q_inspection_result,E1,E2,EG1_export_group)

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Set PQIG030= Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If 

Dim StrData
Dim PvArr
ReDim PvArr(UBound(EG1_export_group,2))
    
For LngRow = 0 To UBound(EG1_export_group,2)

	strData = Chr(11) & ConvSPChars(EG1_export_group(EG1_Insp_Item_Cd, LngRow)) & _
			  Chr(11) & "" & _
			  Chr(11) & ConvSPChars(EG1_export_group(EG1_Insp_Item_Nm, LngRow)) & _
			  Chr(11) & EG1_export_group(EG1_insp_order, LngRow) & _
			  Chr(11) & EG1_export_group(EG1_insp_series, LngRow) & _
			  Chr(11) & UniNumClientFormat(EG1_export_group(EG1_sample_qty, LngRow), ggQty.DecPoint ,0) & _
			  Chr(11) & UniNumClientFormat(EG1_export_group(EG1_accpt_decision_qty, LngRow), ggQty.DecPoint ,0) & _
			  Chr(11) & UniNumClientFormat(EG1_export_group(EG1_rejt_decision_qty, LngRow), ggQty.DecPoint ,0)
			
	If Trim(EG1_export_group(EG1_accpt_decision_discreate, LngRow)) <> "" Then
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(EG1_accpt_decision_discreate, LngRow), 4, 0)
	Else
		strData = strData & Chr(11) & ""
	End If
			
	If Trim(EG1_export_group(EG1_max_defect_ratio, LngRow)) <> "" Then
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(EG1_max_defect_ratio, LngRow), 4, 0)
	Else
		strData = strData & Chr(11) & ""
	End If
			
	strData = strData & Chr(11) & ConvSPChars(EG1_export_group(EG1_InspMethodCd_nm, LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_InspUnitIndctn_nm, LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_insp_spec, LngRow))
			
	If Trim(EG1_export_group(EG1_lsl, LngRow)) <> "" Then
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(EG1_lsl, LngRow), 4, 0)
	Else
		strData = strData & Chr(11) & ""
	End If
			
	If Trim(EG1_export_group(EG1_usl, LngRow)) <> "" Then
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(EG1_usl, LngRow), 4, 0)
	Else
		strData = strData & Chr(11) & ""
	End If
			
	strData = strData & Chr(11) & ConvSPChars(EG1_export_group(EG1_measmt_equipmt_nm, LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_measmt_unit_cd , LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_Insp_Method_Cd , LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_insp_unit_indctn, LngRow)) _
					  & Chr(11) & ConvSPChars(EG1_export_group(EG1_measmt_equipmt_cd, LngRow)) _
					  & Chr(11) & LngMaxRow + LngRow + 1 & Chr(11) & Chr(12)
	PvArr(LngRow) = strData
Next

strData = Join(PvArr, "")
%>  		    
<Script Language="vbscript">    
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		.frm1.vspdData.focus
		.SetAllInspStandardOk
	End with
</Script>
<%
	Set PQIG030 = Nothing
%>
