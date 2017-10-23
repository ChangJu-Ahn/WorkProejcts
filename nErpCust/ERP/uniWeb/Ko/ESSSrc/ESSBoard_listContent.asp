<% Option Explicit %>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
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
