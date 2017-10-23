<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<HTML>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	Response.Expires = -1                               '☜: will expire the response immediately
	Response.Buffer = True                              '☜: The server does not send output to the client until all of the ASP 
														'    scripts on the current page have been processed
'	Dim lgObjConn, lgObjRs
'	Dim lgOpModeCRUD, iStrData
'	Dim lgStrSQL
	Dim nJobTotal, nJobDone
	Dim nRecTotal, nRecDone
	
'	On Error Resume Next
'	Err.Clear
	Call LoadBasisGlobalInf()
	'---------------------------------------Common-----------------------------------------------------------
	Call SubOpenDB(lgObjConn)                           '☜: Make a DB Connection
	lgStrSQL = "SELECT job_id, job_title, job_state, hresult, " & _
			   "	   total_row,  ISNull(succes_row, 0) AS succes_row,  ISNull(failed_row, 0) AS failed_row " & _
			   "FROM   b_bdc_jobs " & _
			   "WHERE  job_id IN ('" & Replace(Request("txtJobs"), " ", "','") & "') " & _
			   "ORDER BY job_id "

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then           'If data not exists
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf
		Response.Write "alert('자료가 없습니다.!');" & vbCrLf
		Response.Write "</SCRIPT>"
	Else
		nJobTotal = 0
		nJobDone = 0
		nRecTotal = 0
		nRecDone = 0
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf

		Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
			If Trim(lgObjRs("hresult")) = "S" Then
				nJobDone = nJobDone + 1
			End If
			
			nRecTotal = lgObjRs("total_row")
			nRecDone = CInt(lgObjRs("succes_row")) + Cint(lgObjRs("failed_row"))

	        Response.Write "parent.nRecTotal[" & nJobTotal & "]=" & nRecTotal & ";" & vbCrLf
			nJobTotal = nJobTotal + 1

			lgObjRs.MoveNext
		Loop

        Response.Write "parent.nJobTotal=" & nJobTotal & ";" & vbCrLf
        Response.Write "parent.nJobDone=" & nJobDone & ";" & vbCrLf
        Response.Write "parent.nRecDone=" & nRecDone & ";" & vbCrLf
        Response.Write "parent.QueryOk();" & vbCrLf
		Response.Write "</SCRIPT>"
	End If

	Call SubCloseRs(lgObjRs)                            '☜ : Release RecordSSet
    Call SubCloseDB(lgObjConn)                          '☜: Close DB Connection
 %>
