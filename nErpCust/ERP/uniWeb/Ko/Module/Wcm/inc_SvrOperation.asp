

<% session.CodePage=949 %>
<%

Const gCLSIDFPMSK	   = """CLSID:9C40F053-0D27-11d2-8810-0000C0E5948C"""
Const C_REVISION_YM = "200703"	' -- 개정판 
Dim wgCO_CD				' -- 사용할 법인 
Dim wgFISC_YEAR
Dim wgCO_NM
Dim wgREP_TYPE
Dim wgGoToPGM 
Dim wgModulePath

wgCO_CD		= "" & Session("CO_CD")
wgCO_NM		= "" & Session("CO_NM")
wgFISC_YEAR = "" & Session("FISC_YEAR")
wgREP_TYPE	= "" & Session("REP_TYPE")
wgGoToPGM	= "" & Session("MNU_ID")
wgModulePath = "" & Session("MODULE_PATH")

If wgCO_CD = "" Then
	If Instr(1, LCase(Request.ServerVariables("SCRIPT_NAME")), "wb121m") = 0 And _
	   Instr(1, LCase(Request.ServerVariables("SCRIPT_NAME")), "wb101m") = 0 And _
	   Instr(1, LCase(Request.ServerVariables("SCRIPT_NAME")), "wb102m") = 0 Then
		If CheckMA_PGM() Then
			Session("MNU_ID") = ReadPGM()	' -- 되돌아가야할 프로그램 
			Call PrintMesg("사용할 법인을 먼저 선택하십시오")
			Call JumpPGM("wb121ma1", "")
			Session("MNU_ID") = ""
		End If
	End If
ElseIf wgModulePath <> "" Then	' 버젼이 다른 경우, 현재 SCRIPT_NAME 에서 모듈을 비교한다.
	Call CheckModulePath
End If

' -- wb121mab.asp 에서 호출 : 
Sub SetCompanyInfo(Byval pCoCd, Byval pCoNm, Byval pFiscYear, Byval pRepType, Byval pRevisionYM)
	Session("CO_CD")		= "" & pCoCd
	Session("CO_NM")		= "" & pCoNm
	Session("FISC_YEAR")	= "" & pFiscYear
	Session("REP_TYPE")		= "" & pRepType
	
	If pRevisionYM <> C_REVISION_YM Then	' 현재 프로그램 버젼과 선택한 법인의 버젼이 다른경우 
		Session("MODULE_PATH")	= "module_" & pRevisionYM
	Else
		Session("MODULE_PATH")	= ""
	End If

	If wgGoToPGM <> "" Then
		Call JumpPGM(wgGoToPGM, ".parent")
	End If
End Sub

' -- 각 mb 단에서 호출 
Function CheckVersion(Byval pFiscYear, Byval pRepType)
	
	Dim sSQL
	
	CheckVersion = False
	
	sSQL = "SELECT REVISION_YM FROM TB_COMPANY_HISTORY WITH (NOLOCK) " & vbCrLf
	sSQL = sSQL & "WHERE CO_CD='" & wgCO_CD & "'"& vbCrLf
	sSQL = sSQL & "	AND	FISC_YEAR='" & pFiscYear & "'"& vbCrLf
	sSQL = sSQL & "	AND REP_TYPE='" & pRepType & "'"& vbCrLf
	
    If   FncOpenRs("R",lgObjConn,lgObjRs,sSQL, "", "") = False Then
  
        Call Displaymsgbox("WC0037", vbInformation, pFiscYear, "년", I_MKSCRIPT)             '☜ : No data is found.
    Else
		' -- 버전 체크 
		If lgObjRs("REVISION_YM") <> C_REVISION_YM Then	' -- 다를때 
			Call Displaymsgbox("WC0035", vbInformation, C_REVISION_YM, lgObjRs("REVISION_YM").value, I_MKSCRIPT) 
			lgObjRs.Close
			Set lgObjRs = Nothing
		Else	' -- 일치할때 
			CheckVersion = True
			Exit Function
		End If
	End If
    
    lgOpModeCRUD = 0	' -- 초기화   
End Function

' -- asp의 파일명을 읽음 
Function ReadPGM()
	Dim pURL, iPos, iPos2
	pURL = Request.ServerVariables("SCRIPT_NAME")
	iPos = Instr(1, pURL, ".asp")
	iPos2 = Instrrev(pURL, "/", iPos-1)
	ReadPGM = Mid(pURL, iPos2+1, iPos-iPos2-1)
End Function

' -- 모듈패스를 읽어서 현재버젼하고 다르면 해당 모듈패스를 추가해 점프함 
Function CheckModulePath()
	Dim sNowModulePath, pURL, sTmp, item, sPath
	
	sNowModulePath = ReadModulePath() ' -- 현재 사용모듈패스 
	
	If wgModulePath <> sNowModulePath And ReadModuleVersion(wgModulePath) <> C_REVISION_YM Then
		' 사용중인 모듈패스와 사용할 법인의 모듈패스가 다른 경우 
		pURL = Replace(LCase(Request.ServerVariables("SCRIPT_NAME")), sNowModulePath, wgModulePath)

		If Server.MapPath(pURL) = "" Then
			pURL = Replace(LCase(Request.ServerVariables("SCRIPT_NAME")), sNowModulePath, "") ' -- 최신버젼은 패스가 엄다 
		End If
		
		sTmp = "?"
		For Each item In Request.QueryString
			sTmp = sTmp & item & "=" & Request.QueryString(item) & "&"
		Next

		pURL = pURL & sTmp
		'Session("MODULE_PATH") = ""	' 제거 2006-01-03 : 법인선택에서 과거선택후 제대로 안되서 
		'Response.Write pURL
		Response.Redirect pURL
		Response.End
	End If
	
End Function

Function ReadModuleVersion(Byval pVer)
	ReadModuleVersion = Replace(LCase(pVer), "module_" , "")	' -- 2006-01-03 : LCase 수정 
End Function

' -- 모듈패스를 읽음 
Function ReadModulePath()
	Dim pURL, iPos, iPos2
	pURL = LCase(Request.ServerVariables("SCRIPT_NAME"))
	iPos = Instr(1, pURL, "/module")
	iPos2 = Instr(iPos+1, pURL, "/")
	ReadModulePath = Mid(pURL, iPos + 1, iPos2 - iPos - 1)
End Function

' -- .asp?뒤에 파라메타를 읽어옴 
Function GetASPParam()
	Dim pURL, iPos, iPos2
	pURL = LCase(Request.ServerVariables("SCRIPT_NAME"))
	iPos = Instr(1, pURL, "/module")
	iPos2 = Instr(iPos+1, pURL, "/")
	ReadModulePath = Mid(pURL, iPos + 1, iPos2 - iPos - 1)
End Function

Function CheckMA_PGM()
	Dim pURL, iPos, sType
	pURL = LCase(Request.ServerVariables("SCRIPT_NAME"))
	iPos = Instr(1, pURL, ".asp")
	sType = Mid(pURL, iPos-3, 2)
	If UCase(sType) = "MA" Then
		CheckMA_PGM = True 
	Else
		CheckMA_PGM = False
	End If
End Function

Function JumpPGM(Byval pPGMID, Byval pParent)
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	Call " & pParent & ".DBGo(""" & pPGMID & """, false)" & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
		Response.End 
End Function

' 메세지 출력 
Sub PrintMesg(Byval strMesg)
%>
<body>
<form name=a><textarea name=txtMesg style="display: none"><%=strMesg%> </textarea></form>
<script language=javascript>
alert(a.txtMesg.value);
</script>
</body>
<%
End Sub


%>