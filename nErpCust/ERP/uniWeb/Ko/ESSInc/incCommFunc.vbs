  Dim gURLPath 
Dim gServerIP
Dim gAmtOfMoney,gEmpNo,gUsrNm,gPws,gDeptAuth,gProAuth,gNumDec,gChildDeptAuth
Call GetGlobalVar_uniSIMS()
'========================================================================================
' Function Name : GetGlobalVar
' Function Desc : Get Global variables from uniConnector
'========================================================================================
Sub GetGlobalVar_uniSIMS()

    On Error Resume Next 
	Set xmlDoc = CreateObject("MSXML2.DOMDocument")	
		
	xmlDoc.async = false 

	xmlDoc.LoadXML GetGlobalXML()
	
	NodeNm = "LoadBasisGlobalInf"

	gADODBConnString	 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gADODBConnString").text
	gAPDateFormat        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateFormat").text
	gAPDateSeperator     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateSeperator").text
	gAPNum1000           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNum1000").text
	gAPNumDec            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNumDec").text
	gAPServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPServer").text
	gClientDateFormat    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateFormat").text	
	gClientDateSeperator = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateSeperator").text	
	gClientNum1000       = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNum1000").text	
	gClientNumDec        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNumDec").text		
	gComDateType		 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComDateType").text	
	gComNum1000			 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNum1000").text
	gComNumDec           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNumDec").text	        
	gConnectionString    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gConnectionString").text
	gDatabase            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDatabase").text
	gDateFormat          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormat").text

	gDateFormatYYYYMM	 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormatYYYYMM").text
	gDBServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServer").text
	gDsnNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDsnNo").text  
	gLocRndPolicy        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLocRndPolicy").text	
	gTaxRndPolicy        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gTaxRndPolicy").text
	gAltNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAltNo").text	     
	gBConfMinorCD        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBConfMinorCD").text
	gCompany             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm &  "/" & "gCompany").text    
	gCompanyNm           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompanyNm").text
	gCurrency            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCurrency").text
	gLang                = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLang").text
	gUsrId               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrId").text
	gRdsUse              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gRdsUse").text
	
	NodeNm ="GetGlobalInf"    
	gDBLoginID           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginID").text
	gDBLoginPwd          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginPwd").text
	gDBServerIP          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerIP").text		
	gDBServerNm          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerNm").text	    
	gUsrNm               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrNm").text
	gEmpNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gEmpNo").text
	gProAuth             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gProAuth").text
	gDeptAuth            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDeptAuth ").text			

    gRowSep  = Chr(12)
    gColSep  = Chr(11)
    NodeNm ="Login"      
	gServerIP			 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "HttpWebSvrIPURL").text  

    gEbEnginePath = "/EasyBaseWeb/bin/AppServer.dll?"
    gEbEnginePath5 = "/REQUBE/bin/RQISAPI.dll?"
    gEbUserName   = "admin"
    gEbDbName     = gDatabase
    gEbPkgRptPath = "Windows\Report"
    gEbPkgRptPath5 = "Report"
    gEbUsrRptPath = "Windows\Report"
	gEbUsrRptPath5 = "Report"
    gEbUserPass   = "admin"

    gEnvInf =           gADODBConnString			&  Chr(12)   '0
    gEnvInf = gEnvInf & gLang						&  Chr(12)   '1
    gEnvInf = gEnvInf & gStrRequestMenuID			&  Chr(12)   '2
    gEnvInf = gEnvInf & gUsrId						&  Chr(12)   '3    
    gEnvInf = gEnvInf & gClientNm					&  Chr(12)   '4
    gEnvInf = gEnvInf & gClientIp					&  Chr(12)   '5    
    gEnvInf = gEnvInf & gUsrId						&  Chr(12)   '6
    gEnvInf = gEnvInf & gSeverity					&  Chr(12)   '7

	Select Case UCase(gLang)
		Case "KO","TEMPLATE"
		         Response.CharSet = "euc-kr"                               'Korea
		         gLogoName = "대사우서비스"
		         gLogo = "ESS"
		Case "CN"
		          Response.CharSet = "GB2312"                               'China
		          gLogoName = "ESS"
		          gLogo = "ESS"
		Case "JA"
		          Response.CharSet = "shift_jis"                            'Japan
		          gLogoName = "ESS"
		          gLogo = "ESS"                
		Case "EN"
		         'Response.CharSet = "windows-1252"                         'U.S.A
		          gLogoName = "ESS"
		          gLogo = "ESS"                
	End Select    
	Set xmlDoc = Nothing
End Sub

Function GetGlobalXML()   '2003-08-07 leejinsoo
    Dim uni2kCommon
    Dim xmlDoc

    On Error Resume Next
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = False 
    Set uni2kCommon    = CreateObject("uni2kCommon.ConnectorControl")
	xmlDoc.LoadXML (uni2kCommon.XivDx(GetSessionStream))
	GetGlobalXML = xmlDoc.xml
    Set xmlDOMDocument = Nothing
    Set uni2kCommon = Nothing
    
End Function

Function GetSessionStream()
    Dim iXmlHttp
    
    On Error Resume Next

    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP")
    iXmlHttp.open "POST", GetIncFolderPath & "/SessionStream.asp", False
    iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    iXmlHttp.send 

    GetSessionStream = iXmlHttp.responseText
  
    Set iXmlHttp = Nothing

End Function


Function GetComaspFolderPath()
   Dim iStrTemp
   Dim iPath   
   Dim i
   
   iStrTemp = Document.Location.href
   iStrTemp = Split(iStrTemp,"/")
   
   For i = 0 To 4
      iPath = iPath & iStrTemp(i) & "/"
   Next
   GetComaspFolderPath = iPath & "Comasp/"
End Function


Function GetIncFolderPath()
   Dim iStrTemp
   Dim iPath   
   Dim i
   
   iStrTemp = Document.Location.href
   iStrTemp = Split(iStrTemp,"/")
   
   For i = 0 To 4
      iPath = iPath & iStrTemp(i) & "/"
   Next
   GetIncFolderPath = iPath & "Inc/"
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Tag disable or visible관련 함수 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'======================================================================================================
'	Function Name	: ProtectTag
'	Description	    : protect object
'======================================================================================================
Sub ProtectTag(objName)

    If UCase(objName.tagName) = "INPUT" Then
       Select Case UCase(objName.TYPE)
          Case "TEXT"
              objName.tabindex = "-1"
'	      objName.className = "protected"
	      objName.className = "form02"
              objName.readonly = True
          Case "CHECKBOX"
              objName.className = "protected"
              objName.disabled = "true"
          Case "RADIO"
              objName.className = "protected"
              objName.disabled = "true"
       End Select
       Exit Sub
    End If

    If UCase(objName.tagName) = "TEXTAREA" Then
       objName.tabindex = "-1"
       objName.className = "form02"
       objName.readonly = True
       Exit Sub
    End If

    If UCase(objName.tagName) = "SELECT" Then
       objName.className = "form02"
       objName.disabled = "true"
       Exit Sub
    End If

'	If objName.type = "text" or objName.tagName = "TEXTAREA" Then
'		If objName.tabindex <> "-1" Then objName.tabindex = "-1"
'		If objName.className <> "protected" Then objName.className = "protected"
'		If objName.readonly <> True Then objName.readonly = True
'	ElseIF objName.type = "radio" or objName.type = "checkbox" or objName.tagName = "SELECT" Then
'		objName.className = "protected"
'		objName.disabled = "true"
'	ElseIf objName.tagname = "OBJECT" Then
'	End If


End Sub

'======================================================================================================
'	Function Name	: ReleaseTag
'	Description	    : Release protected object
'======================================================================================================
Sub ReleaseTag(objName)

    If UCase(objName.tagName) = "INPUT" Then
       Select Case UCase(objName.TYPE)
          Case "TEXT"
             If not isnull(objName.getAttribute("required")) Then
                objName.className = "required"
                objName.readonly = false
                objName.tabindex = ""
             ElseIf not isnull(objName.getAttribute("protected")) Then
                Call ProtectTag(objName)
             ElseIf not isnull(objName.getAttribute("default")) Then
                objName.className = "default"
                objName.readonly = False
                objName.tabindex = ""
             Else
                objName.className = "default"
                objName.readonly = False
                objName.tabindex = ""
             End If
          Case "CHECKBOX", "RADIO"
             If not isnull(objName.getAttribute("required")) Then
                objName.className = "required"
                objName.disabled = "false"
                objName.tabindex = ""
             Else
                objName.className = "default"
                objName.disabled = "false"
                objName.tabindex = ""
             End If
       End Select
       Exit Sub
    End If

    If UCase(objName.tagName) = "TEXTAREA" Then
       If not isnull(objName.getAttribute("required")) Then
          objName.className = "required"
          objName.readonly = false
          objName.tabindex = ""
       ElseIf not isnull(objName.getAttribute("protected")) Then
          Call ProtectTag(objName)
       ElseIf not isnull(objName.getAttribute("default")) Then
          objName.className = "default"
          objName.readonly = False
          objName.tabindex = ""
       Else
          objName.className = "default"
          objName.readonly = False
          objName.tabindex = ""
       End If
       Exit Sub
    End If

    If UCase(objName.tagName) = "SELECT" Then
       If not isnull(objName.getAttribute("required")) Then
          objName.className = "required"
          objName.disabled = "false"
          objName.tabindex = ""
       Else
          objName.className = "default"
          objName.disabled = "false"
          objName.tabindex = ""
       End If
       Exit Sub
    End If

'	If objName.type = "text" or objName.tagName = "TEXTAREA" Then
'		
'		If not isnull(objName.getAttribute("required")) Then
'			If objName.className <> "required" Then objName.className = "required"
'			If objName.readonly <> False Then objName.readonly = false
'			If objName.tabindex <> "" Then objName.tabindex = ""
'		ElseIf not isnull(objName.getAttribute("protected")) Then
'			Call ProtectTag(objName)
'		ElseIf not isnull(objName.getAttribute("default")) Then
'			If objName.className <> "default" Then objName.className = "default"
'			If objName.readonly <> False Then objName.readonly = False
'			If objName.tabindex <> "" Then objName.tabindex = ""
'		Else
'			If objName.className <> "default" Then objName.className = "default"
'			If objName.readonly <> False Then objName.readonly = False
'			If objName.tabindex <> "" Then objName.tabindex = ""
'		End If
'
'	ElseIF objName.type = "radio" or objName.type = "checkbox" or objName.tagName = "SELECT" Then
'		If not isnull(objName.getAttribute("required")) Then
'			If objName.className <> "required" Then objName.className = "required"
'			If objName.disabled <> "false" Then objName.disabled = "false"
'			If objName.tabindex <> "" Then objName.tabindex = ""
'		Else
'			If objName.className <> "default" Then objName.className = "default"
'			If objName.disabled <> "false" Then objName.disabled = "false"
'			If objName.tabindex <> "" Then objName.tabindex = ""
'		End If
'	ElseIf objName.tagname = "OBJECT" Then
'	End If
	
	
End Sub

'========================================================================================
' Sub Name : BtnDisabled(Status)
' Sub Desc : Batch에서의 버튼 활성 / 비활성 상태를 설정한다.
'========================================================================================
Sub BtnDisabled(Status)

	Dim elmCnt, objBtn

	On Error Resume Next

	For elmCnt = 1 to document.body.all.length - 1
	
		Set objBtn = window.document.body.all(elmCnt)
	
		If Ucase(objBtn.TagName) = "BUTTON" then
			objBtn.disabled = Status
		end if
	Next
	
	Set objBtn = Nothing
	
	If Err.Number = 0 Then Err.Clear				    			

End Sub 

'========================================================================================
' Sub Name : elementEnabled(Status)
' Sub Desc : protected된 콤보, 체크, 라디오버튼을 Disable Or Enable한다.
'========================================================================================
Sub ElementEnabled(Status)
	
	Dim elmCnt, objTemp
	
	Status = Not Status

	For elmCnt = 1 to window.document.body.all.length - 1
		Set objTemp = window.document.body.all(elmCnt)
		
		If (Ucase(objTemp.TagName) = "SELECT" Or Ucase(objTemp.TagName) = "RADIO" Or Ucase(objTemp.TagName) = "CHECKBOX") And objTemp.className = "protected" then
			objTemp.disabled = Status
		End If
	Next
	
	Set objCbo = Nothing
	
End Sub

'========================================================================================
' Sub Name : ElementVisible(objElement, Status)
' Sub Desc : Element의 visible속성을 설정한다.
'========================================================================================
Sub ElementVisible(objElement, Status)
	If Status = 0 Then 
		Status = "hidden"
	Else
		Status = "visible"
	End If
	objElement.style.visibility = Status
End Sub

'========================================================================================
' Sub Name : LayerShowHide(Status)
' Sub Desc : 마우스 포인터용 Layer의 Visibility 설정 
'========================================================================================
Sub LayerShowHide(ByVal Status)
	Dim LayerN

	On Error Resume Next
	
	If Status = 0 Then 
		Status = "hidden"
	Else
		Status = "visible"
	End If

	Set LayerN = parent.window.document.all("MousePT").style
	if Err.Number = 0 Then 
		LayerN.visibility = Status
	Else
		Err.Clear				    			
	end if		
End Sub 


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'From POST GET관련 함수 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'========================================================================================
' Function Name : RunMyBizASP
' Function Desc : 비지니스 로직 ASP에 Get 방식으로 실행시킨다.
'========================================================================================
Sub RunMyBizASP(objIFrame, strURL)
	Call BtnDisabled(True)
	strURL = ValueEscape(strURL)
	
	objIFrame.location.href = GetUserPath & strURL
End Sub


'========================================================================================
' Function Name : ExecMyBizASP
' Function Desc : 비지니스 로직 ASP에 Post 방식으로 실행시킨다.
'========================================================================================
Sub ExecMyBizASP(objForm, strAction)
	Call BtnDisabled(True)
	objForm.action = GetUserPath & strAction
	Call elementEnabled(True)
	objForm.submit
	Call elementEnabled(False)
End Sub

'========================================================================================
' Function Name : ConvSPChars
' Function Desc : 문자열안의 "를 ""로 바꾼다.
'========================================================================================
Function ConvSPChars(ByVal strVal)
	ConvSPChars = Replace(strVal, """", """""")
End Function 

'========================================================================================
' Function Name : ValueEscape
' Function Desc : GET 방식으로 넘어가는 문자열에서 name=value 에서 value를 escape 시킨다 
'========================================================================================
Function ValueEscape(strURL)
	Dim szTarget, szAmp
	Dim szValz, szTemp
	Dim arrToken()
	Dim s, e, i, nCnt, s1, e1

	s = Instr(1, strURL, "?")
	If s = 0 Then
		ValueEscape = strURL
		Exit Function
	End If

	szTarget = Left(strURL, s)
	szValz = Mid(strURL, s+1, Len(strURL) - s + 1)

	i = 1
	nCnt = 0
	sp = 1
	Do While Instr(i, szValz, "=") <> 0 
		s = Instr(i, szValz, "=")
		e = Instr(s+1, szValz, "&")
		If e = 0 Then
			e = Len(szValz) + 1
		End If

		s1 = Instr(s+1, szValz, "=")
		e1 = Instr(e+1, szValz, "&")

		If s1 > e1 And e1 <> 0 Then
			szTemp = Mid(szValz, e, s1 - e)
			i = 1
			Do While Instr(i, szTemp, "&") <> 0 
				s1 = Instr(i, szTemp, "&")

				i = s1 + 1
			Loop
			e = e + s1 - 1
		End If

		Redim Preserve arrToken(1, nCnt)
		arrToken(0, nCnt) = Mid(szValz, sp, s - sp)
		arrToken(1, nCnt) = Mid(szValz, s + 1, e - s - 1)
		
		sp = e + 1
		nCnt = nCnt + 1
		i = e + 1
	Loop

	ValueEscape = szTarget 
	szAmp = ""

	For i = 0 To UBound(arrToken, 2)
		If i = 0 Then
			szAmp = ""
		Else
			szAmp = "&"
		End If

'		szTemp = arrToken(1, i)

		arrToken(1, i) = escape(arrToken(1, i))
		
		arrToken(1, i) = Replace(arrToken(1, i), "+", "%2B")
		arrToken(1, i) = Replace(arrToken(1, i), "/", "%2F")

		ValueEscape = ValueEscape + szAmp + arrToken(0, i) + "=" + arrToken(1, i)
		
		
		
'		If arrToken(1, i) <> szTemp Then
'			ValueEscape = ValueEscape + szAmp + arrToken(0, i) + "=" +        arrToken(1, i)
'		Else
'			ValueEscape = ValueEscape + szAmp + arrToken(0, i) + "=" + escape(arrToken(1, i))
'		End If


	Next
End Function
'========================================================================================
' Function Name : GetUserPath
' Function Desc : 현재 디렉토리 패스 알아오기 
'========================================================================================
Function GetUserPath()
	If gURLPath = "" or isEmpty(gURLPath) Then
		Dim strLoc, iPos , iLoc, strPath
		strLoc = window.location.href
                iLoc = inStr(1, strLoc, "?")
            
                If iLoc > 0 Then
                   strLoc = Left(strLoc, iLoc - 1)
                End If
		
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		gURLPath = Left(strLoc, iPos)
	End If
	GetUserPath = gURLPath
End Function

'========================================================================================
' Function Name : GetProgId
' Function Desc : 현재 디렉토리 패스 알아오기 
'========================================================================================
Function GetProgId()

	Dim strLoc, iPos , iLoc, strAspName
	strLoc = UCase(window.location.href)
	
	iLoc = 1: iPos = 0
	
	Do Until iLoc <= 0						
		iLoc = inStr(iPos+1, strLoc, "/")
		If iLoc <> 0 Then iPos = iLoc
	Loop
	strAspName = Right(strLoc, Len(strLoc) - iPos)
	GetProgId = Left(strAspName, Len(strAspName) - Len(".ASP"))
			
    iStr = GetProgId
    If iStr <> "" Then
        iPosTemp = instr(iStr,".")
        If iPosTemp > 0 Then
            iStr = Left(iStr,ipostemp - 1)
        End If
    End If
    GetProgId =iStr 
End Function


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'기타 관련 함수 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'========================================================================
' TAG PARSOR(테그를 인식하여 FORM Element별로 Color부여 
'========================================================================
Function LockField(pDoc)
   On Error Resume Next

    LockField = False
    Dim i, intDivCnt, intTagNum, intCon
    Dim strTagName, strRequired
    Dim iRet
    Dim iRet2
    Dim iRequired,iProtected,iDefault,gProtected,gRequired
    
    iRequired  = UCase(UCN_REQUIRED)
    iProtected = UCase(UCN_PROTECTED)
    iDefault   = UCase(UCN_DEFAULT)
    iTitle     = UCase(UCN_GRID_TITLE)
    gProtected = UCase(UCN_GPROTECTED)
    gRequired  = UCase(UCN_GREQUIRED)
    tProtected = UCase(UCN_TPROTECTED)
    
    intDivCnt = 0
    LockField = False

    For i = 0 To pDoc.All.Length - 1
        strTagName = ""
        intTagNum = 0
        strRequired = ""
        strTagMouseType = ""
        strTagName = UCase(pDoc.All(i).tagName)
        If strTagName <> Empty Then
            If strTagName = "DIV" Then
                intDivCnt = intDivCnt + 1
            End If
        End If

        intCon = Mid(pDoc.All(i).getAttribute("tag"), 1, 1)                
        intTagNum = Mid(pDoc.All(i).getAttribute("tag"), 2, 1)
        strTagMouseType = Mid(pDoc.All(i).getAttribute("tag"), 3, 1)
	    if intTagNum <> "" then
                Select Case strTagName
                    Case "INPUT", "TEXTAREA", "SELECT"
			            if  intTagNum = "4" then
			                if intCon <> "1" then
	                            pDoc.All(i).className = iProtected
	                         end if
                           	if strTagName = "SELECT" OR pDoc.All(i).getAttribute("TYPE") = "checkbox" OR pDoc.All(i).getAttribute("TYPE") = "radio" then
                           	    pDoc.All(i).disabled = true
                           	else
				                pDoc.All(i).readonly = true
				            end if
                		    pDoc.All(i).tabindex = "-1"
			            elseif  intTagNum = "2" then
                           	pDoc.All(i).className = iRequired
                           	if strTagName = "SELECT" then
                           	    pDoc.All(i).disabled = false
                           	else
				                pDoc.All(i).readonly = false
				            end if
                		    pDoc.All(i).tabindex = ""
			            elseif  intTagNum = "3" then
	                        pDoc.All(i).className = iTitle
                           	if strTagName = "SELECT" then
                           	    pDoc.All(i).disabled = true
                           	else
				                pDoc.All(i).readonly = true
				            end if
                		    pDoc.All(i).tabindex = "-1"
			            elseif  intTagNum = "5" then
	                        pDoc.All(i).className = gProtected
                           	if strTagName = "SELECT" then
                           	    pDoc.All(i).disabled = true
                           	else
				                pDoc.All(i).readonly = true
				            end if
                		    pDoc.All(i).tabindex = "-1"
			            elseif  intTagNum = "6" then
	                        pDoc.All(i).className = gRequired
                           	if  strTagName = "SELECT" then
                           	    pDoc.All(i).disabled = false
                           	elseif strTagName = "CHECKBOX" then
                           	        if  pDoc.All(i).value = "Y" then
                           	            pDoc.All(i).className = gProtected
                           	            pDoc.All(i).readonly = true
                           	            pDoc.All(i).disabled = true
                           	         else
                           	         end if
                           	else
				                pDoc.All(i).readonly = false
				            end if
				            pDoc.All(i).tabindex = ""
			            elseif  intTagNum = "9" then
	                        pDoc.All(i).className = tProtected
                           	if strTagName = "SELECT" then
                           	    pDoc.All(i).disabled = true
                           	else
				                pDoc.All(i).readonly = true
				            end if
				            pDoc.All(i).tabindex = ""
			            else
        	            	pDoc.All(i).className = iDefault
				            pDoc.All(i).readonly = false
				            pDoc.All(i).tabindex = ""
			            end if

                        If strTagMouseType="1" Then
                            pDoc.All(i).style.cursor = "hand"
                        End If
                End Select
	    end if
    Next
    LockField = True
End Function

Sub LockElement(Element)
   On Error Resume Next
	Dim intTagNum,strTagMouseType,strTagName

    Dim iRequired,iProtected,iDefault,gProtected,gRequired
    
    iRequired  = UCase(UCN_REQUIRED)
    iProtected = UCase(UCN_PROTECTED)
    iDefault   = UCase(UCN_DEFAULT)
    iTitle     = UCase(UCN_GRID_TITLE)
    gProtected = UCase(UCN_GPROTECTED)
    gRequired  = UCase(UCN_GREQUIRED)

    strTagName = UCase(Element.tagName)
    intTagNum = Mid(Element.getAttribute("tag"), 2, 1)
    strTagMouseType = Mid(Element.getAttribute("tag"), 3, 1)
	if intTagNum <> "" then
        Select Case strTagName
            Case "INPUT", "TEXTAREA", "SELECT"
		        if  intTagNum = "4" then
	                Element.className=iProtected
    	            Element.readonly = true
                    Element.tabindex = "-1"
		        elseif  intTagNum = "2" then
                    Element.className=iRequired
		        	Element.readonly = false
                    Element.tabindex = ""
		        elseif  intTagNum = "3" then
	                Element.className=iTitle
    	            Element.readonly = true
                    Element.tabindex = "-1"
		        elseif  intTagNum = "5" then
	                Element.className=gProtected
    	            Element.readonly = true
                    Element.tabindex = "-1"
		        elseif  intTagNum = "6" then
	                Element.className=gRequired
		        	Element.readonly = false
		        	Element.tabindex = ""
		        else
		        end if
                    If strTagMouseType="1" Then
                        Element.style.cursor = "hand"
                    End If
        End Select
	end if
End Sub


'========================================================================
' 필수입력 항목 체크 
'========================================================================
Function ChkField(pDoc, pStrGrp)
    On Error Resume Next
    ChkField = False
    
    Dim i, intDivCnt, intTagNum, intTagNum3, intTagNum4
    Dim strTagName, strRequired
    Dim iRet
    Dim iRet2
    Dim iRequired
    dim iMaxLen, strNopoint
    dim clNum
    
    iRequired = UCase(UCN_REQUIRED)
    intDivCnt = 0
    ChkField = False
    
    For i = 0 To pDoc.All.Length - 1
        strTagName = ""
        intTagNum = 0
        strRequired = ""
        
        strTagName = UCase(pDoc.All(i).tagName)
        intTagNum =  Mid(pDoc.All(i).getAttribute("tag"), 1, 1)   ' 1: 조회 2:내용 
        intTagNum3 =  Mid(pDoc.All(i).getAttribute("tag"), 3, 1)   ' 숫자포맷 - I : 정수, F : 실수,   날짜  D :년월일  M: 년월 
		intTagNum4 =  Mid(pDoc.All(i).getAttribute("tag"), 4, 1)   ' 부호체크 U : 0포함 양수, S : 음, 0, 양 
		
        If strTagName <> Empty Then
            If strTagName = "DIV" Then
                intDivCnt = intDivCnt + 1
            End If
        End If
                
        strRequired = UCase(pDoc.All(i).className)
        
        If Err.Number <> 0 Then
            Err.Clear
        Else
            'If intTagNum = pStrGrp And strRequired = iRequired Then
'                Select Case strTagName
 '                   Case "INPUT", "TEXTAREA", "SELECT"
  '                      If Len(Trim(pDoc.All(i).Value)) = 0 Then
   '                         If pStrGrp = "1" Then
    '                            pDoc.All(i).focus
     '                           msgbox "조회 필수 항목입니다."
      '                          ChkField = True
       '                         Exit Function
        '                    Else
         '                       pDoc.All(i).focus
          '                      msgbox "입력 필수 항목입니다."
           '                     ChkField = True
            '                    Exit Function
             '               End If
              '          End If
                        
               ' End Select
            'end if
            
            if intTagNum3 = "I" or intTagNum3 = "F" then  '숫자포맷 체크 
            	Select Case strTagName
                    Case "INPUT"
						Select Case num_chk_internal(pDoc.All(i).value, intTagNum3 , intTagNum4)
							Case 1  ' 숫자포맷이 아님 
								Call DisplayMsgBox("229924","X","X","X")
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
							Case 2  ' 부호가 틀림 
            					Call DisplayMsgBox("800484","X",pDoc.All(i).alt ,"X")
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
							Case 3  ' 정수, 실수 구분이 틀림 
            					Call DisplayMsgBox("229924","X","X" ,"X")
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
						End Select
						
						'If UCase(pDoc.All(i).Type) = "TEXT" Then
                         '  iMaxLen = CDbl(pDoc.All(i).maxLength)
                          ' If iMaxLen < 256 Then
                           '   If strRequired <> iProtected Then    
							'	 strNopoint = uniCdbl(pDoc.All(i).value)
								 
                             '    If CmpCharLength(strNoPoint,iMaxLen) = false Then
                              '      iRet = DisplayMsgBox("900028", "X", pDoc.All(i).alt,"x")
                               '     pDoc.All(i).focus
                                '    Set gActiveElement = document.activeElement  
                                 '   ChkField = True                          
                                  '  Exit Function
                        '         End If
                         '     End If
                          ' End If
                        'End If
						
				End select
			elseif intTagNum3 = "D" or intTagNum3 = "M" then
				Select Case strTagName
                    Case "INPUT"
                    Select Case Date_chk_internal(pDoc.All(i).value, intTagNum3)
							Case 1  ' 날짜포맷이 아님 
								Call DisplayMsgBox("174223","X",pDoc.All(i).alt,"X")
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
							Case 2  ' 하한 미만 
            					Call DisplayMsgBox("800504","X",UNIDateClientFormat(gServerBaseDate) ,UNIDateClientFormat(gCommMaximumDate))
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
							Case 3  ' 상한 초과 
            					Call DisplayMsgBox("800504","X",UNIDateClientFormat(gServerBaseDate) ,UNIDateClientFormat(gCommMaximumDate))
								pDoc.All(i).focus()
								ChkField = True
								exit function
								
					End Select
				end select
	        End If
            
        End If
    Next
	ChkField = not ChkFieldLength(pDoc, pStrGrp)    
    
End Function

'=============================================================================
' Function Name  : ClearField
' Parameter      : strString -> Message text
'                  strTarget -> "%"
' Description    : 화면의 그룹별로 데이터를 초기화 합니다.
' Return Value   : "%" Count
'=============================================================================
Function ClearField(pDoc, pStrGrp)
    Err.Clear                                                                    '☜: Clear err status

    For i = 0 To pDoc.All.Length - 1
        strTagName = ""
        intTagNum = 0
        strRequired = ""
        
        strTagName = UCase(pDoc.All(i).tagName)
        If strTagName <> Empty Then
            If strTagName = "DIV" Then
                intDivCnt = intDivCnt + 1
            End If
        End If
                
        intTagNum = Mid(pDoc.All(i).getAttribute("tag"), 1, 1)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If intTagNum=CStr(pStrGrp) Then
            
                Select Case strTagName
                    Case "INPUT", "TEXTAREA", "SELECT"
                		pDoc.All(i).value = ""
                        If pDoc.All(i).style.display="none" Then
                            pDoc.All(i).style.display=""
                        End If
                End Select
            End If
        End If            
    Next
End Function

'=============================================================================
' Function Name  : CountStrings
' Parameter      : strString -> Message text
'                  strTarget -> "%"
' Description    : This function is counting "%" value
' Return Value   : "%" Count
'=============================================================================

Function CountStrings(strString, strTarget)
    Dim lPosition
    Dim lCount
   
    lPosition = 1
    
    Do While InStr(lPosition, strString, strTarget)
    
        lPosition = InStr(lPosition, strString, strTarget) + 1
        lCount = lCount + 1
    
    Loop    
    CountStrings = lCount
   
End Function




'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    if parent.txtDEPT_AUTH.value = "N" then
        msgbox("자료권한이 없습니다.")
        exit function
    end if

    Call MakeKeyStream("N")
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '☜: Direction
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncNext = True                                                               '☜: Processing is OK
	
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status

    if parent.txtDEPT_AUTH.value = "N" then
        msgbox("자료권한이 없습니다.")
        exit function
    end if

    Call MakeKeyStream("P")

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtPrevNext="      & "P"	                         '☜: Direction
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncPrev = True                                                               '☜: Processing is OK
	
End Function

Function Date_chk(pDate, oDate)
    Dim strYear1,strMonth1,strDay1,strDate

    strDate =uniConvDateAtoB(pDate, gDateFormat, gClientDateFormat)
    if Isdate(strDate) and strDate <> "" then
        
        strDate = uniConvDateAtoB(pDate, gDateFormat, gServerDateFormat)
		if strDate <= gServerBaseDate then
			Date_chk = false
		elseif strDate >= gCommMaximumDate then
			Date_chk = false
		else
			oDate = pDate
			Date_chk = true
		end if
    else
        Date_chk = False
    end if
    
End Function



Function Num_chk(pNum, oNum)

    On Error Resume Next
    Err.Clear()
    
    Dim strNum
	
    strNum = unicdbl(trim(pNum))
    
    if Err.number = 0 then   ' cdbl 에서 오류가 생기지 않으면 
    	if  isNumeric(strNum) then
			oNum = strNum
			Num_chk = true
		else
			Num_chk = False
		end if
	else                     ' cdbl에서 오류가 생기면 숫자가 아니다 
		Num_chk = false
	end if
	Err.Clear()
	
End Function



Function Num_chk_internal(byVal pNum, byVal typeTag, byVal signTag)
' parameter : pNum : 숫자 
'             typeTag : "I"     정수 
'                       "F"     실수 
'             signTag : "U"      0, 양수 
'                       "S", "" 음, 0, 양 
' retrun    : 0 정상 
'             1 숫자아님 
'             2 부호이상 
'             3 타입이상 
    On Error Resume Next
    Err.Clear()
    
    Dim strNum
    
    Num_chk_interNal = 0
    
    strNum = unicdbl(trim(pNum))
    
    if Err.number = 0 then   ' cdbl 에서 오류가 생기지 않으면 
		select case signTag
			case "U"
				if strNum  < 0 then
					Num_chk_interNal = 2
					exit function
				end if
			case "S", ""
		end select
		
    	if  isNumeric(strNum) then
    		Select case typeTag
    			case "I" ' 정수 
    				if int(strNum) <> strnum then		
    					Num_chk_interNal = 2
    					exit function
    				end if
    			case "F" ' 실수	
    		End Select
		else
			Num_chk_interNal = 1
		end if
	else                     ' cdbl에서 오류가 생기면 숫자가 아니다 
		Num_chk_interNal = 1
	end if
	Err.Clear()
	
End Function

Function Date_chk_internal(byVal pDate, byVal typeTag)
' parameter : pNum : 숫자 
'             typeTag : "D"     년월일 
'                       "M"     년월 
' retrun    : 0 정상 
'             1 날짜 아님             
'             2 하한미만 
'             3 상한초과 
    On Error Resume Next
    Err.Clear()
    Dim strYear1,strMonth1,strDay1,strDate
	
    strDate =uniConvDateAtoB(pDate, gDateFormat, gClientDateFormat)
    
    Select case typeTag
		case "D"
			if Isdate(strDate) and strDate <> "" then
				strDate = uniConvDateAtoB(pDate, gDateFormat, gServerDateFormat)
				if strDate < gServerBaseDate then
					Date_chk_internal = 2
				elseif strDate > gCommMaximumDate then
					Date_chk_internal = 3
				else
					Date_chk_internal = 0
				end if
			else
				Date_chk_internal = 1
			end if
		Case "M"
			call ExtractDateFrom(pDate,gDateFormatYYYYMM , gComDateType   ,strYear1,strMonth1,strDay1)
			strDate = strYear1 & gComDateType & strMonth1 & gComDateType & "01"
			strDate = uniConvDateAtoB(strDate, gServerDateFormat, gDateFormat)
			if Isdate(strDate) and strDate <> "" then
				strDate = uniConvDateAtoB(strDate, gDateFormat, gServerDateFormat)
				if strDate < gServerBaseDate then
					Date_chk_internal = 2
				elseif strDate > gCommMaximumDate then
					Date_chk_internal = 3
				else
					Date_chk_internal = 0
				end if
			else
				Date_chk_internal = 1
			end if
	End Select
	
	Err.Clear()
	
End Function


'========================================================================================================
' Name : OpenCalendar()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCalendar(InID, pParm)
	If document.activeElement.className = "protected" Then Exit Function
	Dim arrRet
	Dim arrParam(2)
	Dim InObj
	Dim strDate

    Set InObj = document.all(InID)
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = InObj.value
	
	arrRet = window.showModalDialog("./Calendar.asp", InObj.value, _
		"dialogWidth=350px; dialogHeight=226px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

    If arrRet = "" Then
		Exit Function
    End IF

	strDate = arrRet

	Select Case pParm
		Case 1
		    InObj.value = year(strDate)
		Case 2
		    InObj.value = UNIFormatMonth(UNICDate(strDate))
		Case 3
		    InObj.value = strDate
	End Select
    Set InObj = nothing
		
End Function

'======================================================================================================
' 설명 : FGetFormatDate
' 기능 : 
'======================================================================================================
Function FGetFormatDate(Byval InDate, Byval InCase)
Dim RDate
Dim strDt
Dim strTt
Dim strSt 

    If InDate <> "" Then
        InDate = FormatDateTime(InDate,2)
        If instr(InDate,"-") then
            strDt = Year(InDate) & Mid(InDate,instr(InDate,"-"))
        End If
        strTt = FormatDateTime(InDate,4)
        strSt = FormatDateTime(InDate,3)
        If instr(strSt,":") Then
            strSt = Mid(strSt,instrRev(strSt,":"))
        End If
        Rdate = strDt & Chr(32) & Chr(32) & Chr(32) & strTt & strSt
    End If

    Select Case Ucase(InCase)
        Case "Y"
            RDate = Year(InDate)
        Case "M"
            RDate = Month(InDate)
        Case "D"
            RDate = Day(InDate)
        Case "T"
            RDate = Time(InDate)
        Case "YMDT"
            Rdate = Rdate
        Case "YMD"
            Rdate = strDt
        Case "YM"
            Rdate = Mid(Rdate, 1, 7)            
        Case "MD"
            Rdate = Mid(Rdate, 6, 10)        
    End Select

    FGetFormatDate = RDate
End Function    

Function SetToolBar(inpar)

    ' 조회 
    if  mid(inpar,1,1) = "1" then
	    parent.submit.style.display = ""
	Else
	    parent.submit.style.display = "none"
    end if

    ' 삭제 
    if  mid(inpar,3,1) = "1" then
	    parent.del.style.display = ""
	Else
	    parent.del.style.display = "none"
    end if

    ' 추가 
    if  mid(inpar,2,1) = "1" then
	    parent.add.style.display = ""
	Else
	    parent.add.style.display = "none"
    end if

    ' 저장 
    if  mid(inpar,4,1) = "1" then
	    parent.save.style.display = ""
	Else
	    parent.save.style.display = "none"
    end if

    ' 출력 
    if  mid(inpar,5,1) = "1" then
	    parent.prt.style.display = ""
	Else
	    parent.prt.style.display = "none"
    end if

End Function
'========================================================================================
' Function Name : UNIMsgBox
' Function Desc : 
'========================================================================================
Function UNIMsgBox(pVal, pType, pTitle)
	MsgBox pVal, pType, pTitle
End Function
