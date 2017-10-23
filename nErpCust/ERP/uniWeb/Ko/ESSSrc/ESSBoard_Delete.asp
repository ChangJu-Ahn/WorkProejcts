<%	Option Explicit%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
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
