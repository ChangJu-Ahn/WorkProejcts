<%@ Language=VBScript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%

	Response.Expires=0
	Response.Clear 
	Response.ContentType = "text/xml"
	
    Call LoadBasisGlobalInf()  
    'Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

	Dim sCoCd, sFiscYear, sRepType, sUsrID, sMnuID
    sCoCd		= FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    sFiscYear	= FilterVar(Request("txtFISC_YEAR"),"''", "S")	' 사업연도 
    sRepType	= FilterVar(Request("cboREP_TYPE"),"''", "S")		' 신고구분 
    sUsrID		= FilterVar(gUsrID,"''", "S")		' 신고구분 
	sMnuID		= Request("txtMNU_ID")		' 메뉴ID
	
	Const adExecuteStream = 1024
	Const MSSQLXML_DIALECT = "{5D531CB2-E6Ed-11D2-B252-00C04F681B71}"
	
	Dim conDB 'As ADODB.Connection
	Dim cmdXML 'As ADODB.Command
	Dim stmXMLout 'AS ADODB.Stream
	Dim strQry 'As String
	Dim gConnect, sLang, sXML, sSQL
	
	Set conDB = CreateObject("ADODB.Connection")

	' Connect to the database using Integrated Security.
	With conDB
	    '.Provider = "SQLOLEDB"
	    .ConnectionString = gADODBConnString
	    .Open

	End With
	
	Set cmdXML = CreateObject("ADODB.Command")

	'Assign the Connection object to the Command object.
	Set cmdXML.ActiveConnection = conDB

	'Create the query template.
	sSQL = ""
	strQry = "<BMNU xmlns:sql='urn:schemas-microsoft-com:xml-sql'>" & vbCrLf
	strQry = strQry & "<sql:query>"		 & vbCrLf
	strQry = strQry & "EXEC dbo.usp_TB_TAX_DOC_CheckProgress " & sCoCd & "," & sFiscYear & "," & sRepType & "," & sUsrID & vbCrLf
	sSQL = sSQL & "SELECT A.* ,  CASE WHEN CONFIRM_FLG = '1' THEN '3' ELSE ISNULL(D.STATUS_FLG, '') END STATUS_FLG" & vbCrLf
	sSQL = sSQL & "FROM V_MENU	A" & vbCrLf
	sSQL = sSQL & "	INNER JOIN TB_TAX_DOC_DTL D ON A.CALLED_FRM_ID = D.PGM_ID AND D.CO_CD=" & sCoCd & " AND D.FISC_YEAR=" & sFiscYear & " AND D.REP_TYPE = " & sRepType  & vbCrLf
	sSQL = sSQL & "WHERE A.USR_ID = " & sUsrID & " " & vbCrLf
	sSQL = sSQL & "AND A.UPPER_MNU_ID = '" & sMnuID & "' " & vbCrLf
	sSQL = sSQL & "ORDER BY UPPER_MNU_ID, MNU_SEQ" & vbCrLf
	strQry = strQry & sSQL
	strQry = strQry & "FOR XML RAW" & vbCrLf
	strQry = strQry & "</sql:query>" & vbCrLf
	strQry = strQry & "<Debug><![CDATA[" & sCoCd & "," & sFiscYear & "," & sRepType & "," & gUsrId & "]]></Debug>" & vbCrLf
	strQry = strQry & "<DebugSQL><![CDATA[" & sSQL & "]]></DebugSQL>" & vbCrLf
	strQry = strQry & "</BMNU>" & vbCrLf

	'Specify the MSSQLXML dialect and assign the query.
	cmdXML.Dialect = MSSQLXML_DIALECT
	cmdXML.CommandText = strQry

	'Create Stream object for results.

	Set stmXMLout = CreateObject("ADODB.Stream")

	'Assign the result stream.
	stmXMLout.Open
	cmdXML.Properties("Output Stream") = stmXMLout

	'Execute the query.
	cmdXML.Execute , , adExecuteStream

	
	Select Case UCASE(glang)
	Case "KO","TEMPLATE","TEMPLATE1"
	    sLang  = "euc-kr"                               'Korea
	Case "CN"
	    sLang = "GB2312"                               'China
	Case "JA"
	    sLang = "shift_jis"                            'Japan
	Case "EN"
	    sLang = "euc-kr"
	   'Response.CharSet = "windows-1252"                         'U.S.A
	Case "HU"
	    sLang = "windows-1250"                         'Hungary
	End Select   

	sXML = "<?xml version='1.0' encoding='euc-kr'?>" & vbCrLf 
	sXML = sXML & stmXMLout.ReadText
	
	Set conDB = Nothing
	Set cmdXML = Nothing
	Set stmXMLout = Nothing
	
	Response.Write sXML
   
%>


	