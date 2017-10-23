<% Option Explicit %>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
	Dim seqs, page,from_where,SearchPart,SearchStr
	dim table:table = "ESS_Board"
	seqs = Request("seqs")
	page = Request("page")
	SearchPart = Request("SearchPart")
	SearchStr = Request("SearchStr")	

	from_where = Request("from_where")  
	if right(seqs,1) = "," then seqs = left(seqs, len(seqs)-1)
  
    Call SubOpenDB(lgObjConn)  

    lgStrSQL = "Update " & table & " Set readCount=readCount+1 " & " WHERE seq  in ( " & seqs & ")"	

	call	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")
	if from_where = "" then
		Response.Redirect "ESSBoard_content.asp?" & "seqs=" & seqs & "&page=" & page
	else
		Response.Redirect "ESSBoard_content.asp?" & "seqs=" & seqs & "&page=" & page& "&from_where=s&SearchPart=" &SearchPart&"&SearchStr="&SearchStr
	end if	
%>
