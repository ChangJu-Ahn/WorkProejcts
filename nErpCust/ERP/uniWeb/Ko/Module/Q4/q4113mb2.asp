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
'*  3. Program ID           : Q4113MB2
'*  4. Program Name         : 수입검사불합격통지 
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

'IMPORTS VIEW
Const Q340_I1_frame_dt = 0
Const Q340_I1_framer = 1
Const Q340_I1_defect_comment = 2
Const Q340_I1_defect_contents = 3
Const Q340_I1_required_improvement = 4

Dim lgIntFlgMode
Dim iStrSelectChar
Dim PQIG340
Dim iImportRejectReport
Dim iStrPlantCd
Dim iStrInspReqNo

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
	
If Len(Trim(Request("txtFrameDt"))) Then
	If UNIConvDate(Request("txtFrameDt")) = "" Then
		Call DisplayMsgBox("122116", vbinformation, "", "", I_MKSCRIPT)
		Response.End
	End If
End If

If lgIntFlgMode = OPMD_CMODE Then
	iStrSelectChar = "CREATE"
	iStrPlantCd = Request("txtPlantCd")	
	iStrInspReqNo = Request("txtInspReqNo2")
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iStrSelectChar = "UPDATE"
	iStrPlantCd = Request("hPlantCd")	
	iStrInspReqNo = Request("txtInspReqNo1")
End If

Redim iImportRejectReport(4)
iImportRejectReport(Q340_I1_frame_dt) = UNIConvDate(Request("txtFrameDt"))
iImportRejectReport(Q340_I1_framer) = Request("txtFramer")
iImportRejectReport(Q340_I1_defect_comment) = Request("txtDefectComment")
iImportRejectReport(Q340_I1_defect_contents) = Request("txtDefectContents")
iImportRejectReport(Q340_I1_required_improvement) = Request("txtRequiredImprovement")

Set PQIG340 = Server.CreateObject("PQIG340.cQMtRejReportSimple")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG340.Q_MAINT_REJECT_REPORT_SIMPLE_SVR(gStrGlobalCollection, _
									iStrSelectChar, _
									iStrPlantCd, _
									iStrInspReqNo, _
									1, _
									iImportRejectReport)
	    
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG340 = Nothing
	Response.End
End If		    
		              
Set PQIG340 = Nothing   
%>
<Script Language=vbscript>
	Call Parent.DbSaveOk()
</Script>