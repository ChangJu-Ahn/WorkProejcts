<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" --> <!--uniDateClientFormate 있을때 만 쓰고 없으면 뺀다.-->
<!-- #Include file="../../inc/IncSvrNumber.inc" --> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %> <!--숫자형 있을때 만 쓰고 없으면 뺀다.-->
<%Call LoadBasisGlobalInf%>
<%
'********************************************************************************************************
'*  1. Module Name          :Quality											*
'*  2. Function Name        : 
'*  3. Program ID           : q2611pb1.asp												*
'*  4. Program Name         : q24118ListOccurSvr 														*
'*  5. Program Desc         : 
'*  7. Modified date(First) : 2000/05/09					
'*  8. Modified date(Last)  : 2000/05/09					
'*  9. Modifier (First)     : Koh jae woo
'* 10. Modifier (Last)      : Koh jae woo									*
'* 11. Comment              :																			*
'********************************************************************************************************

On Error Resume Next

Call HideStatusWnd 
													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim PQIG210

Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strMgmtNo
Dim strWcCd
Dim strFrameDt1
Dim strFrameDt2
Dim strCounterPlanFlag
Dim StrData
Dim i
Dim lgStrPrevKey
Dim NextKey

Const C_SHEETMAXROWS_D = 100
		
		
Dim I3_q_assignable_occurrence
ReDim I3_q_assignable_occurrence(5)
'[CONVERSION INFORMATION]  IMPORTS View 상수 
Const Q407_I3_mgmt_no = 0    '[CONVERSION INFORMATION]  View Name : import q_assignable_occurrence
Const Q407_I3_insp_class_cd = 1
Const Q407_I3_wc_cd = 2
Const Q407_I3_plant_cd = 3
Const Q407_I3_item_cd = 4
Const Q407_I3_counter_plan_flag = 5


Dim EG1_group_export
ReDim EG1_group_export(15)
'[CONVERSION INFORMATION]  EXPORTS Group View 상수 
'[CONVERSION INFORMATION]  Group Name : group_export
Const Q407_EG1_E1_mgmt_no = 0    '[CONVERSION INFORMATION]  View Name : export q_assignable_occurrence
Const Q407_EG1_E1_insp_class_cd = 1
Const Q407_EG1_E1_frame_dt = 2
Const Q407_EG1_E1_occur_dt_fr = 3
Const Q407_EG1_E1_occur_dt_to = 4
Const Q407_EG1_E1_wc_cd = 5
Const Q407_EG1_E1_plant_cd = 6
Const Q407_EG1_E1_item_cd = 7
Const Q407_EG1_E1_contents_of_assignable_occur = 8
Const Q407_EG1_E1_reason_for_occur = 9
Const Q407_EG1_E1_framer = 10
Const Q407_EG1_E1_counter_plan_flag = 11
Const Q407_EG1_E2_plant_nm = 12    '[CONVERSION INFORMATION]  View Name : export b_plant
Const Q407_EG1_E3_item_nm = 13    '[CONVERSION INFORMATION]  View Name : export b_item
Const Q407_EG1_E4_wc_nm = 14    '[CONVERSION INFORMATION]  View Name : export p_work_center
Const Q407_EG1_E5_minor_nm = 15    '[CONVERSION INFORMATION]  View Name : export_nm_for_insp_class_cd b_minor
		
strPlantCd 	= Request("txtPlantCd")
strItemCd 	= Request("txtItemCd")
strInspClassCd= Request("txtInspClassCd")
strMgmtNo 	= Request("txtMgmtNo")
strWcCd	 	= Request("txtWcCd")
strFrameDt1	= Request("txtFrameDt1")
strFrameDt2	= Request("txtFrameDt2")
strCounterPlanFlag	= Request("txtCounterPlanFlag")

lgStrPrevKey 	= Request("lgStrPrevKey")
LngMaxRow 	= Request("txtMaxRows")

If strFrameDt1 <> "" then
	strFrameDt1 	= UNIConvDate(strFrameDt1)
End If

If strFrameDt2 <> "" then
	strFrameDt2 	= UNIConvDate(strFrameDt2)
End If

If lgStrPrevKey = "" then
	NextKey	= strMgmtNo
Else
	NextKey	= lgStrPrevKey
End If

I3_q_assignable_occurrence(Q407_I3_mgmt_no) =	NextKey
I3_q_assignable_occurrence(Q407_I3_insp_class_cd) = strInspClassCd
I3_q_assignable_occurrence(Q407_I3_wc_cd) = strWcCd
I3_q_assignable_occurrence(Q407_I3_plant_cd) = strPlantCd
I3_q_assignable_occurrence(Q407_I3_item_cd) = strItemCd
I3_q_assignable_occurrence(Q407_I3_counter_plan_flag) =  strCounterPlanFlag

Set PQIG210 = Server.CreateObject("PQIG210.cQListOccurSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG210.Q_LIST_OCCUR_SVR(gStrGlobalCollection, _
  							  C_SHEETMAXROWS_D, _
			 				  strFrameDt1, _
							  strFrameDt2, _
							  I3_q_assignable_occurrence, _
							  EG1_group_export)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG210 = Nothing
	Response.End
End If


For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_mgmt_no)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_item_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E3_item_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_wc_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E4_wc_nm)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(EG1_group_export(i, Q407_EG1_E1_frame_dt)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_counter_plan_flag)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(EG1_group_export(i, Q407_EG1_E1_occur_dt_fr)))
		strData = strData & Chr(11) & UNIDateClientFormat(Trim(EG1_group_export(i, Q407_EG1_E1_occur_dt_to)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_framer)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_plant_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E2_plant_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E1_insp_class_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q407_EG1_E5_minor_nm)))
		strData = strData & Chr(11) & LngMaxRow + i
		strData = strData & Chr(11) & Chr(12)
    Else
		StrNextKey = EG1_group_export(i, Q407_EG1_E1_mgmt_no)
    End If
Next  

Set PQIG210 = Nothing

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
			 
		 .DbQueryOk
    End If	
End with
</Script>