<%	Option Explicit%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
	Dim seq, page
	dim table:table = "ESS_Board"
	seq = Request.QueryString("seq")
	page = Request.QueryString("page")
	
'	Dim userid	
'	userid = gEmpNo

 '   if userid <> id then
 '		Response.Redirect "Content.asp?" & Request.ServerVariables("QUERY_STRING")
 '		Response.End
  '  end if
	Call SubOpenDB(lgObjConn)  
	lgStrSQL = "delete " & table & " where seq=" & seq

	if	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") =true then
		Call SubCloseDB(lgObjConn)
		Response.Redirect "ESSBoard_list.asp?page=" & page
	else
	end if    
%>
