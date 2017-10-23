<%	Option Explicit%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
	Dim seq, page, ipAddr, subject, Content
	dim table:table = "ESS_Board"
	page = Request.Form("page")
	seq = Request.Form("seq")
	subject = Request.Form("subject")
	if subject <> "" and Len(subject) > 70 then
		subject = left(subject,60) & "..."
	end if
	Content = Request.Form("Content")
	ipAddr = Request.ServerVariables("REMOTE_ADDR")
	
	Call SubOpenDB(lgObjConn)  
	lgStrSQL = "Update "& table & " set ipAddr = "& FilterVar(ipAddr,"''", "S") & ",Subject ="& FilterVar(Subject,"''", "S")
	lgStrSQL = lgStrSQL &  ",Content =" & FilterVar(Content,"''", "S")
	lgStrSQL = lgStrSQL & " where seq=" & seq

	if	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") =true then
		Call SubCloseDB(lgObjConn)
		Response.Redirect "ESSBoard_content.asp?seqs=" & seq & "&page=" & page
	else
	end if
%>
