<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2612PB1
'*  4. Program Name         : 
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
													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim PQIG240

Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strMgmtNo
Dim strWcCd
Dim strFrameDt1
Dim strFrameDt2
Dim strPlanDt1
Dim strPlanDt2
Dim strCounterPlanFlag 
Dim StrData
Dim i

Dim lgStrPrevKey

Const C_SHEETMAXROWS_D = 100

strPlantCd 	= Request("txtPlantCd")
strItemCd 	= Request("txtItemCd")
strInspClassCd	= Request("txtInspClassCd")
strMgmtNo 	= Request("txtMgmtNo")

strWcCd	 	= Request("txtWcCd")
strFrameDt1	= Request("txtFrameDt1")
strFrameDt2	= Request("txtFrameDt2")
strPlanDt1	= Request("txtPlanDt1")
strPlanDt2	= Request("txtPlanDt2")
strCounterPlanFlag	= Request("txtCounterPlanFlag")

lgStrPrevKey 	= Request("lgStrPrevKey")
LngMaxRow 	= Request("txtMaxRows")

''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim I1_q_assignable_occurrence
ReDim I1_q_assignable_occurrence(5)
Const Q415_I1_mgmt_no = 0
Const Q415_I1_insp_class_cd = 1
Const Q415_I1_wc_cd = 2
Const Q415_I1_plant_cd = 3
Const Q415_I1_item_cd = 4
Const Q415_I1_counter_plan_flag = 5
If lgStrPrevKey = "" then
	I1_q_assignable_occurrence(Q415_I1_mgmt_no)	= strMgmtNo
Else
	I1_q_assignable_occurrence(Q415_I1_mgmt_no)	= lgStrPrevKey
End If
I1_q_assignable_occurrence(Q415_I1_insp_class_cd) = strInspClassCd
I1_q_assignable_occurrence(Q415_I1_wc_cd) = strWcCd
I1_q_assignable_occurrence(Q415_I1_plant_cd) = strPlantCd
I1_q_assignable_occurrence(Q415_I1_item_cd) = strItemCd
I1_q_assignable_occurrence(Q415_I1_counter_plan_flag) = strCounterPlanFlag
'''''''''''''''''''''''''''''''''''''''''''''''''''' 
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim I2_q_assignable_occurrence_fr
ReDim I2_q_assignable_occurrence_fr(1)
Const Q415_I2_occur_dt = 0
Const Q415_I2_frame_dt = 1
I2_q_assignable_occurrence_fr(Q415_I2_occur_dt) = ""
If strFrameDt1 <> "" then
	I2_q_assignable_occurrence_fr(Q415_I2_frame_dt) 	= UNIConvDate(strFrameDt1)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim I3_q_assignable_occurrence_to
ReDim I3_q_assignable_occurrence_to(1)
Const Q415_I3_occur_dt = 0
Const Q415_I3_frame_dt = 1
I3_q_assignable_occurrence_to(Q415_I3_frame_dt) = ""
If strFrameDt2 <> "" then
	I3_q_assignable_occurrence_to(Q415_I3_frame_dt) 	= UNIConvDate(strFrameDt2)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim I4_q_assignable_occurrence_result_plan_dt
Dim I5_q_assignable_occurrence_result_plan_dt

If strPlanDt1 <> "" then
	I4_q_assignable_occurrence_result_plan_dt 	= UNIConvDate(strPlanDt1)
End If
If strPlanDt2 <> "" then
	I5_q_assignable_occurrence_result_plan_dt 	= UNIConvDate(strPlanDt2)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''

Set PQIG240 = Server.CreateObject("PQIG240.cQListAssignOccurrRst")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Dim EG1_group_export
ReDim EG1_group_export(17)
Const Q415_EG1_E1_mgmt_no = 0
Const Q415_EG1_E1_insp_class_cd = 1
Const Q415_EG1_E1_frame_dt = 2
Const Q415_EG1_E1_occur_dt_fr = 3
Const Q415_EG1_E1_occur_dt_to = 4
Const Q415_EG1_E1_wc_cd = 5
Const Q415_EG1_E1_plant_cd = 6
Const Q415_EG1_E1_item_cd = 7
Const Q415_EG1_E1_contents_of_assignable_occur = 8
Const Q415_EG1_E1_reason_for_occur = 9
Const Q415_EG1_E1_framer = 10
Const Q415_EG1_E1_counter_plan_flag = 11    
Const Q415_EG1_E1_plant_nm = 12
Const Q415_EG1_E1_result_counter_plan_dt = 13
Const Q415_EG1_E1_result_framer = 14
Const Q415_EG1_E1_result_dtls_of_counter_plan_contents = 15
Const Q415_EG1_E1_item_nm = 16
Const Q415_EG1_E1_wc_nm = 17

Call PQIG240.Q_LIST_ASSIGN_OCCURR_RESULT(gStrGlobalCollection, _
  							  C_SHEETMAXROWS_D, _
			 				  I1_q_assignable_occurrence, _
							  I2_q_assignable_occurrence_fr, _
							  I3_q_assignable_occurrence_to, _
							  I4_q_assignable_occurrence_result_plan_dt, _
							  I5_q_assignable_occurrence_result_plan_dt, _
							  EG1_group_export)
         

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG240 = Nothing
	Response.End
End If


For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_mgmt_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_item_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_item_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_wc_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_wc_nm)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_frame_dt))))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_result_counter_plan_dt))))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_counter_plan_flag)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_occur_dt_fr))))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_occur_dt_to))))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q415_EG1_E1_framer)))
		strData = strData & Chr(11) & LngMaxRow + i
		strData = strData & Chr(11) & Chr(12)
    Else
		StrNextKey = EG1_group_export(i, Q407_EG1_E1_mgmt_no)
    End If
Next  

Set PQIG240 = Nothing
%>		    
<Script Language="vbscript">
With parent
	.ggoSpread.Source = .vspdData 
	.ggoSpread.SSShowData "<%=strData%>"
	.vspdData.focus
		
	.lgStrPrevKey = "<%=StrNextKey%>"
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> "" Then
		 .DbQuery
	Else
		 <% ' Request값을 hidden Varialbe로 넘겨줌 %>
		 .hPlantCd 	= "<%=ConvSPChars(strPlantCd)%>"
		 .hItemCd	= "<%=ConvSPChars(strItemCd)%>"
		 .hMgmtNo	= "<%=ConvSPChars(strMgmtNo)%>"
		 .hInspClassCd = "<%=ConvSPChars(strInspClassCd)%>"
		 .hWcCd		= "<%=ConvSPChars(strWcCd)%>"
		 .hFrameDt1	= "<%=strFrameDt1%>"
		 .hFrameDt2	= "<%=strFrameDt2%>"
		 .hPlanDt1	= "<%=strPlanDt1%>"
		 .hPlanDt2	= "<%=strPlanDt2%>"
			 
		 .DbQueryOk
    End If
End with
</Script>