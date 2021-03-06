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
'*  3. Program ID           : Q2612MB1
'*  4. Program Name         : 공정이상대책 보고서 등록 
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

Dim PQIG230
Dim strMgmtNo
Dim E1_q_assignable_occurrence
Redim E1_q_assignable_occurrence(11)
Const Q413_E1_mgmt_no = 0
Const Q413_E1_insp_class_cd = 1
Const Q413_E1_frame_dt = 2
Const Q413_E1_occur_dt_fr = 3
Const Q413_E1_occur_dt_to = 4
Const Q413_E1_wc_cd = 5
Const Q413_E1_plant_cd = 6
Const Q413_E1_item_cd = 7
Const Q413_E1_contents_of_assignable_occur = 8
Const Q413_E1_reason_for_occur = 9
Const Q413_E1_framer = 10
Const Q413_E1_counter_plan_flag = 11	
    
Dim E2_p_work_center
Redim E2_p_work_center(0)
const Q413_E2_wc_nm = 0    

Dim E3_b_plant
Redim E3_b_plant(0)	
const Q413_E3_plant_nm = 0

Dim E4_b_item
Redim E4_b_item(0)
const Q413_E4_item_nm = 0

Dim E5_q_assignable_occurrence_result
Redim E5_q_assignable_occurrence_result(2)
Const Q413_E5_counter_plan_dt = 0
Const Q413_E5_framer = 1
Const Q413_E5_dtls_of_counter_plan_contents = 2		
		
	
strMgmtNo = Request("txtMgmtNo")

Const C_SHEETMAXROWS_D = 100

Set PQIG230 = Server.CreateObject("PQIG230.cQLookupOccurRstSvr")


If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG230.Q_LOOK_UP_OCCUR_RESULT_SVR (gStrGlobalCollection , _
										C_SHEETMAXROWS_D, _
										strMgmtNo , _
										E1_q_assignable_occurrence , _
										E2_p_work_center, _
										E3_b_plant , _
										E4_b_item, _
										E5_q_assignable_occurrence_result)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG230 = Nothing	
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG230 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	.txtMgmtNo2.Value = "<%=ConvSPChars(strMgmtNo)%>"
	.txtPlantCd.Value =  "<%=ConvSPChars(E1_q_assignable_occurrence(Q413_E1_plant_cd))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(E3_b_plant(Q413_E3_plant_nm))%>"
	.cboInspClassCd.Value = "<%=ConvSPChars(E1_q_assignable_occurrence(Q413_E1_insp_class_cd))%>"
	.txtFramer.Value = "<%=ConvSPChars(E1_q_assignable_occurrence(Q413_E1_framer))%>"
	.txtItemCd.Value = "<%=ConvSPChars(E1_q_assignable_occurrence(Q413_E1_item_cd))%>"
	.txtItemNm.Value = "<%=ConvSPChars(E4_b_item(Q413_E4_item_nm))%>"
	.txtWcCd.Value = "<%=ConvSPChars(E1_q_assignable_occurrence(Q413_E1_wc_cd))%>"
	.txtWcNm.Value = "<%=ConvSPChars(E2_p_work_center(Q413_E2_wc_nm))%>"
	.txtOccurDtFr.Text = "<%=UNIDateClientFormat(E1_q_assignable_occurrence(Q413_E1_occur_dt_fr))%>"
	.txtOccurDtTo.Text = "<%=UNIDateClientFormat(E1_q_assignable_occurrence(Q413_E1_occur_dt_to))%>"
	.txtFrameDt.Value = "<%=UNIDateClientFormat(E1_q_assignable_occurrence(Q413_E1_frame_dt))%>"
	.txtCounterPlanDt.Text = "<%=UNIDateClientFormat(E5_q_assignable_occurrence_result(Q413_E5_counter_plan_dt))%>"
	.txtCounterPlanFramer.Value = "<%=ConvSPChars(E5_q_assignable_occurrence_result(Q413_E5_framer))%>"
	.txtDtlsOfCounterPlanContents.Value = "<%=ConvSPChars(E5_q_assignable_occurrence_result(Q413_E5_dtls_of_counter_plan_contents))%>"
		
End with
	
parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음		
	
Parent.DbQueryOk
</Script>
<%
Set PQIG230 = Nothing    
%>
  