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
'*  3. Program ID           : Q2611MB3
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
	
Dim PQIG190																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMgmtNo
Dim strInspClassCd
Dim strPlantCd
	
strMgmtNo = Request("txtMgmtNo")	

dim I1_q_assignable_occurrence
Redim I1_q_assignable_occurrence(11)
Const Q401_I1_mgmt_no = 0
Const Q401_I1_insp_class_cd = 1
Const Q401_I1_frame_dt = 2
Const Q401_I1_occur_dt_fr = 3
Const Q401_I1_occur_dt_to = 4
Const Q401_I1_wc_cd = 5
Const Q401_I1_plant_cd = 6
Const Q401_I1_item_cd = 7
Const Q401_I1_contents_of_assignable_occur = 8
Const Q401_I1_reason_for_occur = 9
Const Q401_I1_framer = 10
Const Q401_I1_insrt_user_id = 11

I1_q_assignable_occurrence(Q401_I1_mgmt_no) = strMgmtNo
I1_q_assignable_occurrence(Q401_I1_insp_class_cd) = strInspClassCd
I1_q_assignable_occurrence(Q401_I1_frame_dt) = UNIConvDate(Request("txtFrameDt"))
I1_q_assignable_occurrence(Q401_I1_occur_dt_fr) = UNIConvDate(Request("txtOccurDtFr"))
I1_q_assignable_occurrence(Q401_I1_occur_dt_to) = UNIConvDate(Request("txtOccurDtTo"))
I1_q_assignable_occurrence(Q401_I1_wc_cd) =  Request("txtWcCd")
I1_q_assignable_occurrence(Q401_I1_plant_cd) = strPlantCd
I1_q_assignable_occurrence(Q401_I1_item_cd) = Request("txtItemCd")
I1_q_assignable_occurrence(Q401_I1_contents_of_assignable_occur) = Request("txtContentsofAssignableOccur")
I1_q_assignable_occurrence(Q401_I1_reason_for_occur) = Request("txtReasonForOccur")	
I1_q_assignable_occurrence(Q401_I1_framer) = Request("txtFramer")	
I1_q_assignable_occurrence(Q401_I1_insrt_user_id) = ""

dim I2_ief_supplied
I2_ief_supplied  = "D"	
	
Set PQIG190 = Server.CreateObject("PQIG190.cQMaintOccurSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG190.Q_MAINT_OCCUR_SVR(gStrGlobalCollection, _
							   I1_q_assignable_occurrence, _
						 	   I2_ief_supplied)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG190 = Nothing
	Response.End
End If

Set PQIG190 = Nothing                                                   '☜: Unload Comproxy
%>
<Script Language=vbscript>
Call parent.DbDeleteOk()
</Script>