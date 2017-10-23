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
'*  3. Program ID           : Q2612MB3
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
	
Dim PQIG220																'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMgmtNo
Dim lgIntFlgMode
Dim strUserId
Dim strInspClassCd
Dim strPlantCd
DIM I3_ief_supplied
DIM I2_q_assignable_occurrence_result
	
REDIM I2_q_assignable_occurrence_result(2)
Const Q409_I2_counter_plan_dt = 0
Const Q409_I2_framer = 1
Const Q409_I2_dtls_of_counter_plan_contents = 2
    
strMgmtNo = Request("txtMgmtNo")
strPlantCd = Request("txtPlantCd")
lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
strUserId	= Request("txtInsrtUserId")
strInspClassCd = Request("cboInspClassCd")

If Len(Trim(Request("txtCounterPlanDt"))) Then
	If UNIConvDate(Request("txtCounterPlanDt")) = "" Then
		Call DisplayMsgBox("122116", vbinformation, "", "", I_MKSCRIPT)
		Response.End
	End If
End If
				
I2_q_assignable_occurrence_result (Q409_I2_counter_plan_dt) = Request("txtCounterPlanDt")
I2_q_assignable_occurrence_result(Q409_I2_framer) =  Request("txtCounterPlanFramer")
I2_q_assignable_occurrence_result(Q409_I2_dtls_of_counter_plan_contents) = Request("txtDtlsOfCounterPlanContents")
			
Set PQIG220 = Server.CreateObject ("PQIG220.cQMtOccurResultSvr")
			
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
    	
I3_ief_supplied  = "D"	
   	
Call PQIG220.Q_MAINT_OCCUR_RESULT_SVR  (gStrGlobalCollection, strMgmtNo, _
										I2_q_assignable_occurrence_result, I3_ief_supplied)

	
Set PQIG220 = Nothing      
	
'-----------------------
'Com action result check area(DB,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
              
Set PQIG220 = Nothing                                                  '☜: Unload Comproxy
%>
<Script Language=vbscript>
With parent
	.frm1.txtPlantCd.Value = "<%=ConvSPChars(strPlantCd)%>"											'☜: 화면 처리 ASP 를 지칭함 
	.frm1.txtMgmtNo1.Value = "<%=ConvSPChars(strMgmtNo2)%>"
	.frm1.cboInspClassCd.Value = "<%=ConvSPChars(strInspClassCd)%>"
	
	.DbDeleteOk
End With
</Script>
<%
Set PQIG220 = Nothing                                                   '☜: Unload Comproxy
%>