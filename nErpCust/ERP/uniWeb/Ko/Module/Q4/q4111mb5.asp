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
'*  3. Program ID           : Q4111MB5
'*  4. Program Name         : 검사결과등록 
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

Dim PQIG310																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Const I1_plant_cd = 0
Const I1_insp_req_no = 1
Const I1_insp_result_no = 2
Const I1_release_dt = 3
Const I1_sl_cd_for_good = 4
Const I1_sl_cd_for_defect = 5
	
Dim IG1_q_inspection_result
	
Set PQIG310 = Server.CreateObject("PQIG310.cQMtReleaseSimple")

If CheckSystemError(Err,True) Then					
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
	
Redim IG1_q_inspection_result(5)

IG1_q_inspection_result(I1_plant_cd) = UCase(Request("txtPlantCd"))	    
IG1_q_inspection_result(I1_insp_req_no) = UCase(Request("txtInspReqNo"))
IG1_q_inspection_result(I1_insp_result_no) = 1

IG1_q_inspection_result(I1_release_dt) = UNIConvDate(Request("txtReleaseDt"))
IG1_q_inspection_result(I1_sl_cd_for_good) = UCase(Request("txtGoodsSLCd"))
IG1_q_inspection_result(I1_sl_cd_for_defect) = UCase(Request("txtDefectivesSLCd"))

Call PQIG310.Q_MAINT_INSP_RELEASE_SIMPLE_SVR(gstrGlobalCollection, "CANCEL", IG1_q_inspection_result)

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Set PQIG310 = Nothing
%>
<Script Language=vbscript>
	Parent.DbCancelReleaseOK
</Script>
<%	
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
Set PQIG310 = Nothing                                                  '☜: Unload Comproxy
%>
<Script Language=vbscript>
	Parent.DbCancelReleaseOk
</Script>