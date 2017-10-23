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
'*  3. Program ID           : Q2611MB1
'*  4. Program Name         : 이상발생 보고서 정보등록 
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

Dim PQIG200													'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMgmtNo

Dim E1_b_item_item_nm
Dim E2_b_plant_plant_nm
Dim E3_p_work_center_wc_nm

Dim E4_q_assignable_occurrence
ReDim E4_q_assignable_occurrence(11)

Const Q405_E4_mgmt_no = 0
Const Q405_E4_insp_class_cd = 1
Const Q405_E4_frame_dt = 2
Const Q405_E4_occur_dt_fr = 3
Const Q405_E4_occur_dt_to = 4
Const Q405_E4_wc_cd = 5
Const Q405_E4_plant_cd = 6
Const Q405_E4_item_cd = 7
Const Q405_E4_contents_of_assignable_occur = 8
Const Q405_E4_reason_for_occur = 9
Const Q405_E4_framer = 10
Const Q405_E4_counter_plan_flag = 11
	
strMgmtNo = Request("txtMgmtNo")

Set PQIG200 = Server.CreateObject("PQIG200.cQLookupOccurSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG200.Q_LOOK_UP_OCCUR_SVR(gStrGlobalCollection, _
                                 strMgmtNo, _
                                 E1_b_item_item_nm, _
                                 E2_b_plant_plant_nm, _
                                 E3_p_work_center_wc_nm, _
                                 E4_q_assignable_occurrence)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG200 = Nothing
	Response.End
End If

Set PQIG200 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	.txtPlantCd.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_plant_cd))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(E2_b_plant_plant_nm)%>"
	.txtMgmtNo2.Value = "<%=ConvSPChars(strMgmtNo)%>"
	.cboInspClassCd.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_insp_class_cd))%>"
	.txtItemCd.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_item_cd))%>"
	.txtItemNm.Value = "<%=ConvSPChars(E1_b_item_item_nm)%>"
	.txtWcCd.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_wc_cd))%>"
	.txtWcNm.Value = "<%=ConvSPChars(E3_p_work_center_wc_nm)%>"
	.txtFramer.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_framer))%>"
	.txtOccurDtFr.Text = "<%=UNIDateClientFormat(E4_q_assignable_occurrence(Q405_E4_occur_dt_fr))%>"
	.txtOccurDtTo.Text = "<%=UNIDateClientFormat(E4_q_assignable_occurrence(Q405_E4_occur_dt_to))%>"
	.txtFrameDt.Text = "<%=UNIDateClientFormat(E4_q_assignable_occurrence(Q405_E4_frame_dt))%>"
	.txtContentsofAssignableOccur.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_contents_of_assignable_occur))%>"
	.txtReasonForOccur.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_reason_for_occur))%>"
	.cboCounterPlanFlag.Value = "<%=ConvSPChars(E4_q_assignable_occurrence(Q405_E4_counter_plan_flag))%>"
End with
	
'parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
'parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음		
	
Parent.DbQueryOk
</Script>	