<%	Option Explicit%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%
	dim table:table = "ESS_Board"
	Dim page : page = request("page")
		
	Dim id, name, ipAddr, subject, pwd,content
	name =  Request.Form("name")
	subject = Request.Form("subject")
	if subject <> "" and Len(subject) > 70 then
		subject = left(subject,60) & "..."
	end if
		
	content = Request.Form("content")
	
	if len(content) > 20000 then
		Response.Write "<Script language=javascript>"
		Response.Write "alert('너무 많은 데이터(2만자이상)를 입력하셨습니다');"
		Response.Write "</Script>"
		Response.End
	end if
	
	
'	id = Request.Cookies("N")("userid")
	ipAddr = Request.ServerVariables("REMOTE_ADDR")

	Dim Content_short
	Content_short = Mid(content, 1, 200)
	
	Call SubOpenDB(lgObjConn)  
	if gUsrNm="" and gEmpno="unierp" then
		gUsrNm="admin"
	end if
	
	lgStrSQL = "Insert into "& table & " (id, name, subject,  ipAddr, content,readcount,inputDate) values("
	lgStrSQL = lgStrSQL & FilterVar(gEmpno,"''", "S") & "," & FilterVar(gUsrNm,"''", "S") & "," & FilterVar(subject,"''", "S")  & "," & FilterVar(ipAddr,"''", "S")& ","  & FilterVar(content,"''", "S") & ",0,getdate())"

	if	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") =true then
		Call SubCloseDB(lgObjConn)
		Response.Redirect "ESSBoard_list.asp?page=1"
	end if
%>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   