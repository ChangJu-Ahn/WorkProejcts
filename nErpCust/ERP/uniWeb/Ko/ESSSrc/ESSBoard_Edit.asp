<%	Option Explicit%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
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
