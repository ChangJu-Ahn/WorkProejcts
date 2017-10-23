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
'*  3. Program ID           : Q2414MB3
'*  4. Program Name         : 불량원인등록 
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

Dim strinsp_class_cd
strinsp_class_cd = "P"	'@@@주의 

Dim LngRow
'***** START
Dim LngMaxRow
'***** END
Dim intGroupCount 
Dim PQIG180
Dim lgStrPrevKeyM
Dim lglngHiddenRows
Dim lRow
Dim StrNextKey
Dim strInspReqNo
Dim strInspItemCd
Dim strPlantCd
Dim strInspSeries
Dim strDefectTypeCd
Dim strDefectCause
Dim StrData
Dim i
Dim E1_q_inspection_defect_cause

Dim EG1_export_group
Redim EG1_export_group(1)
Const Q276_EG1_E1_q_inspection_defect_cause_defect_cause_cd = 0
Const Q276_EG1_E1_q_inspection_defect_cause_defect_qty = 1
Dim EG2_export_group
Redim EG2_export_group(0)
Const Q276_EG2_E1_q_defect_cause_defect_cause_nm = 0

Dim I3_q_inspection_details
Redim I3_q_inspection_details(1)
Const Q275_I3_insp_item_cd = 0
Const Q275_I3_insp_series = 1
    
Dim I2_q_inspection_result
Redim I2_q_inspection_result(2)
Const Q276_I2_insp_result_no = 0
Const Q276_I2_insp_class_cd = 1
Const Q276_I2_plant_cd = 2

Const C_SHEETMAXROWS_D = 100

strInspReqNo	= Request("txtInspReqNo")
strInspItemCd	= Request("txtInspItemCd")
strInspSeries	= Request("txtInspSeries")
strPlantCd		= Request("txtPlantCd")
strDefectTypeCd = Request("txtDefectTypeCd")
lgStrPrevKeyM	= Request("lgStrPrevKeyM")
lglngHiddenRows = CLng(Request("lglngHiddenRows"))
lRow			= CLng(Request("lRow"))
'***** START
LngMaxRow = Request("txtMaxRows")
'***** END

I3_q_inspection_details(Q275_I3_insp_item_cd) = strInspItemCd
I3_q_inspection_details(Q275_I3_insp_series) = strInspSeries

I2_q_inspection_result(Q276_I2_insp_result_no) = 1
I2_q_inspection_result(Q276_I2_insp_class_cd) = strinsp_class_cd
I2_q_inspection_result(Q276_I2_plant_cd) =  strPlantCd

If lgStrPrevKeyM <> "" then
	strDefectCause = lgStrPrevKeyM
end If

Set PQIG180 = Server.CreateObject("PQIG180.cQListInspDefCausSvr")

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PQIG180.Q_LIST_INSP_DEFT_CAUSE_SVR (gStrGlobalCollection, C_SHEETMAXROWS_D, strInspReqNo , _
										I2_q_inspection_result, I3_q_inspection_details, strDefectTypeCd, _
										strDefectCause, EG1_export_group, EG2_export_group, E1_q_inspection_defect_cause)
			
If CheckSystemError(Err,True) Then										'☜: ComProxy Unload
	Set PQIG180= Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If


Set PQIG180= Nothing

'Export To Spread2
strData = ""	
intGroupCount = UBound(EG1_export_group, 1)
If intGroupCount < C_SHEETMAXROWS_D Then
	intGroupCount = UBound(EG1_export_group, 1) + 1
End If

For i = 0 To UBound(EG1_export_group,1)
	If i < C_SHEETMAXROWS_D Then
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_export_group(i, Q276_EG1_E1_q_inspection_defect_cause_defect_cause_cd)))
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & Trim(ConvSPChars(EG2_export_group(i, Q276_EG2_E1_q_defect_cause_defect_cause_nm)))
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(i, Q276_EG1_E1_q_inspection_defect_cause_defect_qty), ggQty.DecPoint ,0)
		strData = strData & Chr(11) & lRow 
		strData = strData & Chr(11) & lglngHiddenRows + i + 1 		'Flag
		strData = strData & Chr(11) & LngMaxRow + i + 1
		strData = strData & Chr(11) & Chr(12)			
	Else		
		StrNextKey = EG1_export_group(i,0)
	End If
Next

%>		    
<Script Language="vbscript">   
With Parent
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowDataByClip "<%=strData%>"	
	.lgStrPrevKeyM(<%=lRow-1%>) = "<%=ConvSPChars(StrNextKey)%>"				
	.lglngHiddenRows(<%=lRow-1%>) = "<%=lglngHiddenRows + intGroupCount%>"
	If .frm1.vspdData2.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData2, 0) And .lgStrPrevKeyM(<%=lRow-1%>) <> "" Then
		Call .DbQuery2(<%=lRow%>, True)
	Else
		 <% ' Request값을 hidden input으로 넘겨줌 %>
		 .frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
		 .frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		 .frm1.hInspItemCd.value = "<%=ConvSPChars(strInspItemCd)%>"
		 .frm1.hInspSeries.value = "<%=strInspSeries%>"
		 .frm1.hDefectTypeCd.value = "<%=ConvSPChars(strDefectTypeCd)%>"
		 Call .DbQueryOk2("<%=intGroupCount%>")
    End If		
End with
</Script>
<%
Set PQIG180 = Nothing  
%>