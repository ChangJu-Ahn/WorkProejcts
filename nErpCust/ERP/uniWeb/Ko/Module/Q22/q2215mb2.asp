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
'*  3. Program ID           : Q2215MB2
'*  4. Program Name         : 판정 
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
														
Dim PQIG100																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim lgIntFlgMode
Dim LngMaxRow
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strInspReqNo
Dim LngRow
Dim lGrpCnt
	
Dim I3_q_inspection_result
Redim I3_q_inspection_result(6)
Const Q236_I3_insp_result_no = 0
Const Q236_I3_decision = 1
Const Q236_I3_insp_dt = 2
Const Q236_I3_insp_qty = 3
Const Q236_I3_defect_qty = 4
Const Q236_I3_inspector_cd = 5
Const Q236_I3_rmk = 6
    
Dim IG1_import_Group
Dim iCommand
    
lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
LngMaxRow = CInt(Request("txtMaxRows"))	

Set PQIG100 = Server.CreateObject("PQIG100.cQMtDecisionSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
	
'Header Import
icommand = "CONFIRM"
	
strInspReqNo = Trim(Request("txtInspReqNo"))
	
I3_q_inspection_result(Q236_I3_insp_result_no) = 1
I3_q_inspection_result(Q236_I3_decision) = Request("cboDecision")
If Request("txtInspDt") <> "" Then
	I3_q_inspection_result(Q236_I3_insp_dt) = UniConvDate(Request("txtInspDt"))
End If
I3_q_inspection_result(Q236_I3_insp_qty) = UNIConvNum(Request("txtInspQty"), 0)
I3_q_inspection_result(Q236_I3_defect_qty) = UNIConvNum(Request("txtDefectQty"), 0)
I3_q_inspection_result(Q236_I3_inspector_cd) = Request("txtInspectorCd")
I3_q_inspection_result(Q236_I3_rmk) = Request("txtRemark")

'Detail Import
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	LngMaxRow = Ubound(arrRowVal)
	Redim IG1_import_Group(LngMaxRow,3)
	
	For LngRow = 0 To LngMaxRow - 1
		arrColVal = Split(arrRowVal(LngRow), gColSep)
		lGrpCnt = lGrpCnt + 1						'☜: Group Count
		IG1_import_Group(LngRow,0) = arrColVal(1)
		IG1_import_Group(LngRow,1) = UniConvNum(arrColVal(2),0)
		IG1_import_Group(LngRow,2) = arrColVal(4)
		IG1_import_Group(LngRow,3) = UniConvNum(arrColVal(3),0)
	Next
		
	Call PQIG100.Q_MAINT_DECISION_SVR(gstrglobalcollection, _
	                                  icommand, strInspReqNo, _
	                                  I3_q_inspection_result, _
	                                  IG1_import_Group)
		
	If CheckSYSTEMError(Err,True) Then
		Set PQIG100 = Nothing 
		Response.End
	End If
End If	
Set PQIG100 = Nothing                                                   '☜: Unload Comproxy
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>