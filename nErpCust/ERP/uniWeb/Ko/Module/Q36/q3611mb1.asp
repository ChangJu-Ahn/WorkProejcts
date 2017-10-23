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
'*  1. Module Name          : Quality
'*  2. Function Name        : 
'*  3. Program ID           : Q3611MB1
'*  4. Program Name         : 월집계/확정/확정취소 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/09/04
'*  8. Modified date(Last)  : 2003/09/04
'*  9. Modifier (First)     : Jaewoo Koh
'* 10. Modifier (Last)      : Jaewoo Koh
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
									
On Error Resume Next
Call HideStatusWnd
	
Dim PQIG390		
Dim iStrAction

iStrAction = Request("txtAction")

Set PQIG390 = Server.CreateObject("PQIG390.cQMtSummary")
If CheckSystemError(Err,True) Then					
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

SELECT CASE iStrAction
	CASE "S"		'Summary
		Call PQIG390.Q_MAINT_SUMMARY_PROCESSING_SVR(gstrGlobalCollection, _
											Request("txtPlantCd"), _
											Request("txtYr"), _
											Request("txtMnth"), _
											Request("txtRYesorNo"), _
											Request("txtPYesorNo"), _
											Request("txtFYesorNo"), _
											Request("txtSYesorNo"))			
	CASE "C"		'Confirm
		Call PQIG390.Q_SUMMARY_CONFIRM_SVR(gstrGlobalCollection, _
											Request("txtPlantCd"), _
											Request("txtYr"), _
											Request("txtMnth"), _
											Request("txtRYesorNo"), _
											Request("txtPYesorNo"), _
											Request("txtFYesorNo"), _
											Request("txtSYesorNo"))
			
	CASE "R"		'Cancel Confirm
		Call PQIG390.Q_CANCEL_SUMMARY_CONFIRM_SVR(gstrGlobalCollection, _
											Request("txtPlantCd"), _
											Request("txtYr"), _
											Request("txtMnth"), _
											Request("txtRYesorNo"), _
											Request("txtPYesorNo"), _
											Request("txtFYesorNo"), _
											Request("txtSYesorNo"))
			
END SELECT

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Set PQIG390 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

SELECT CASE iStrAction
	CASE "S"		'Summary
		Call DisplayMsgBox("224710", vbinformation, "", "", I_MKSCRIPT)
	CASE "C"		'Confirm
		Call DisplayMsgBox("224723", vbinformation, "", "", I_MKSCRIPT)
	CASE "R"		'Cancel Confirm
		Call DisplayMsgBox("224724", vbinformation, "", "", I_MKSCRIPT)	
END SELECT

Set PQIG390 = Nothing
Response.End														'☜: 비지니스 로직 처리를 종료함 
%>