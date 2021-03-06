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
'*  3. Program ID           : Q2612MB2
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
	
Dim PQIG220																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Dim lgIntFlgMode
Dim strInspClassCd
Dim strPlantCd
Dim strMgmtNo2
Dim I3_ief_supplied
Dim I2_q_assignable_occurrence_result
	
ReDim I2_q_assignable_occurrence_result(2)
Const Q409_I2_counter_plan_dt = 0
Const Q409_I2_framer = 1
Const Q409_I2_dtls_of_counter_plan_contents = 2
    
	
strMgmtNo2 = Request("txtMgmtNo2")
strPlantCd = Request("txtPlantCd")
lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
strInspClassCd = Request("cboInspClassCd")
	
If Len(Trim(Request("txtCounterPlanDt"))) Then
	If UNIConvDate(Request("txtCounterPlanDt")) = "" Then
		Call DisplayMsgBox("122116", vbinformation, "", "", I_MKSCRIPT)
		Response.End
	End If
End If
		
I2_q_assignable_occurrence_result (Q409_I2_counter_plan_dt) = UNIConvDate(Request("txtCounterPlanDt"))
I2_q_assignable_occurrence_result(Q409_I2_framer) =  Request("txtCounterPlanFramer")
I2_q_assignable_occurrence_result(Q409_I2_dtls_of_counter_plan_contents) = Request("txtDtlsOfCounterPlanContents")
	
Set PQIG220 = Server.CreateObject ("PQIG220.cQMtOccurResultSvr")
	
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If


If lgIntFlgMode = OPMD_CMODE Then
	I3_ief_supplied  = "C"
Else
	I3_ief_supplied  = "U"					
End If
			
Call PQIG220.Q_MAINT_OCCUR_RESULT_SVR  (gStrGlobalCollection, strMgmtNo2, _
										I2_q_assignable_occurrence_result, I3_ief_supplied)
		
		
Set PQIG220 = Nothing      
	
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
%>
<Script Language=vbscript>
With parent
	.frm1.txtPlantCd.Value = "<%=ConvSPChars(strPlantCd)%>"											'☜: 화면 처리 ASP 를 지칭함 
	.frm1.txtMgmtNo1.Value = "<%=ConvSPChars(strMgmtNo2)%>"
	.frm1.cboInspClassCd.Value = "<%=ConvSPChars(strInspClassCd)%>"
	
	.DbSaveOk
End With
</Script>
<%
Set PQIG220 = Nothing                                                   '☜: Unload Comproxy
%>  