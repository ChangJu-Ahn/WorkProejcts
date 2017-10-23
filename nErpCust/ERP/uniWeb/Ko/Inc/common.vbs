Dim gURLPath 
Dim gIMGPath

'========================================================================================
Function MXD(ByVal pData)
    Dim iuni2kCommon
    
    Set iuni2kCommon = CreateObject("uni2kCommon.ConnectorControl")
    
    If Err.Number <> 0 Then
       MsgBox "암호화 모듈 생성 에러 입니다." 
       Exit Function
    End If   
    
    MXD = iuni2kCommon.xCVTG(pData)

    Set iuni2kCommon = Nothing

End Function

Function AndProc(val1, val2)
	Dim i
	
	AndProc = ""
	
	val2 = Left(val2 & String(Len(val1), "0"), Len(val1))

	For i = 1 To Len(val1)
		If Mid(val1, i, 1) = "1" And Mid(val2, i, 1) = "1" Then
			AndProc = AndProc & "1"
		Else
			AndProc = AndProc & "0"
		End If
	Next

End Function

'========================================================================================
Function CheckOCXQuery()
   
	CheckOCXQuery = False
	
    If MSIEVer() = "5.5" Then
       CheckOCXQuery = True
       Exit Function
    End If
	
End Function

'======================================================================================================
Sub DisableToolBar(pOpt)
    Dim iBit
    
    On Error Resume Next
    
    Call RestoreToolBar()

    gActionStatus = pOpt
    
    Select Case pOpt
       Case TBC_QUERY      : iBit = "1011111111111111"
       Case TBC_NEW        : iBit = "1101111111111111"
       Case TBC_DELETE     : iBit = "1110111111111111"
       Case TBC_SAVE       : iBit = "1111011111111111"
       Case TBC_INSERTROW  : iBit = "1111101111111111"
       Case TBC_DELETEROW  : iBit = "1111110111111111"
       Case TBC_CANCEL     : iBit = "1111111011111111"
       Case TBC_PREV       : iBit = "1111111101111111"
       Case TBC_NEXT       : iBit = "1111111110111111"
       Case TBC_COPYRECORD : iBit = "1111111111011111"
       Case TBC_EXPORT     : iBit = "1111111111101111"
       Case TBC_PRINT      : iBit = "1111111111110111"
       Case TBC_FIND       : iBit = "1111111111111011"
       Case TBC_HELP       : iBit = "1111111111111101"
       Case TBC_EXIT       : iBit = "1111111111111110"
       Case Else           : Exit Sub
    End Select    
    gToolBarBit = AndProc(gToolBarBit, iBit)
    Call parent.SetToolbar(gToolBarBit)
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
Sub GetGlobalVar()
       
	Dim iNode 
    Dim FileNm
	Dim xmlDoc
	Dim NodeNm
	Dim objConn
	
	Set objConn = CreateObject("uniConnector.cGlobal")
	objConn.CheckURL(GetURLLangUserID())
		
	set xmlDoc = CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = false 
	
	xmlDoc.LoadXML(objConn("GlobalXMLData2"))
	
	NodeNm = "LoadBasisGlobalInf"
	
	gADODBConnString     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gADODBConnString").text
	gAPDateFormat        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateFormat").text
    gAPDateSeperator     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateSeperator").text
    gAPNum1000           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNum1000").text
    gAPNumDec            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNumDec").text
	gAPServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPServer").text
	gClientDateFormat    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateFormat").text	
    gClientDateSeperator = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateSeperator").text	
    gClientNum1000       = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNum1000").text	
    gClientNumDec        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNumDec").text	
    gColSep              = Chr(11)                    
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
    gRowSep			     = Chr(12)
    gTaxRndPolicy        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gTaxRndPolicy").text       
    gAltNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAltNo").text	     
    gBConfMinorCD        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBConfMinorCD").text
    gCompany             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompany").text    
    gCompanyNm           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompanyNm").text
    gCurrency            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCurrency").text
    gLang                = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLang").text
    gPlant               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPlant").text
    gPlantNm             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPlantNm").text
    gSetupMod            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSetupMod").text
    gSeverity            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSeverity").text
    gUsrId               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrId").text
    gDBKind              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBKind").text
    gRdsUse              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gRdsUse").text      '2003-12-13 leejinsoo
    gUserIdKind          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUserIdKind").text   '2003-05-31 leejinsoo
    
    NodeNm = "GetGlobalInf"    
    
    gBizArea             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBizArea").text
    gBizUnit             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBizUnit").text
    gChangeOrgId         = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gChangeOrgId").text
    gCostCenter          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCostCd").text
    gCountry             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCountry").text
    gDepart              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDepart").text
    gDBLoginID           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginID").text
    gDBLoginPwd          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginPwd").text
    gDBServerIP          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerIP").text		
    gDBServerNm          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerNm").text	    
    gFiscCnt             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscCnt").text
    gFiscEnd             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscEnd").text
    gFiscStart           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscStart").text
    gIntDeptCd           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gIntDeptCd").text
    gIm_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gIm_Post_Flag").text
    gLoginDt             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLoginDt").text    
    gLogonGp             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLogonGp").text
    gPurOrg              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPurOrg").text
    gPurGrp              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPurGrp").text
    gSalesGrp            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSalesGrp").text
    gSalesOrg            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSalesOrg").text
    gSo_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSo_Post_Flag").text
    gPo_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPo_Post_Flag").text
    gStorageLoc          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gStorageLoc").text
    gUsrEngName          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrEngName").text
    gUsrNm               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrNm").text
    gWorkCenter          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gWorkCenter").text
    gClientIp            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientIp").text
    gClientNm            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNm").text

' 2003-02-19 Kim In Tae	
'1: zigzag(right align)    2: decimal point align
	gQMDPAlignOpt = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gQMDPAlignOpt").text
	gIMDPAlignOpt = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gIMDPAlignOpt").text
        
    gEbEnginePath = "/EasyBaseWeb/bin/AppServer.dll?"
    gEbEnginePath5 = "/REQUBE/bin/RQISAPI.dll?"
    gEbUserName   = "admin"
    gEbDbName     = gDatabase
    gEbPkgRptPath = "Windows\Report"
    gEbPkgRptPath5 = "Report"
    gEbUsrRptPath = "Windows\Report"
	gEbUsrRptPath5 = "Report"
    gEbUserPass   = "admin"

    gEnvInf =           gADODBConnString  & Chr(12)   '0
    gEnvInf = gEnvInf & gLang             & Chr(12)   '1
    gEnvInf = gEnvInf & gStrRequestMenuID & Chr(12)   '2
    gEnvInf = gEnvInf & gUsrId            & Chr(12)   '3    
    gEnvInf = gEnvInf & gClientNm         & Chr(12)   '4
    gEnvInf = gEnvInf & gClientIp         & Chr(12)   '5    
    gEnvInf = gEnvInf & gUsrId            & Chr(12)   '6
    gEnvInf = gEnvInf & gSeverity         & Chr(12)   '7
    gEnvInf = gEnvInf & gDBKind           & Chr(12)   '7

    
    set xmlDoc  = nothing
    set objConn = nothing
    
End Sub

'========================================================================================
Function GetImgPath(pVal)
	If gIMGPath = "" or isEmpty(gIMGPath) Then
		Dim strLoc, iPos , iLoc, strPath
		strLoc = pVal
                iLoc = inStr(1, strLoc, "?")
            
                If iLoc > 0 Then
                   strLoc = Left(strLoc, iLoc - 1)
                End If
		
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		gIMGPath = Left(strLoc, iPos)
	End If
	GetImgPath = gIMGPath
End Function

'========================================================================================
Function GetXmlFilePath()
	Dim i
    Dim strCookie   
    Dim strXMLFile
    Dim XMLFile
    Dim HexCookie
	Dim HexToString
	Dim XMLFilePath

	strCookie = Split(document.cookie,"&")
	
		
	for i = 0 to UBound(strCookie)
		if  instr(strCookie(i),"gHTTPXMLFileNm") <> 0  then	
			
			strXMLFile = mid(strCookie(i), (instr(1,strCookie(i),"=")+1))					
       		Exit For   
		end if 
	next		
				
	strXMLFile = Split(strXMLFile, ";")
	XMLFile = strXMLFile(0)
	HexCookie = Split(XMLFile, "%")
	XMLFilePath = HexCookie(0)
	For i = 1 To UBound(HexCookie)
		HexToString = "&H" & Mid(HexCookie(i), 1, 2)
		XMLFilePath = XMLFilePath & Chr(CLng(HexToString)) & Mid(HexCookie(i), 3, Len(HexCookie(i)))
	Next


	XMLFilePath = Replace(XMLFilePath, "+", " ")
	
	GetXmlFilePath =  XMLFilePath
	
End Function

'========================================================================================
' Function Name : GetURLLangUserID
' Function Desc : 현재 URL과 언어 코드,사용자 ID 알아오기 
'========================================================================================
Function GetURLLangUserID()
	Dim i
    Dim strCookie   
    Dim strURLLangUserID
    Dim HexCookie
	Dim HexToString
	Dim URLLangUsrID
	Dim URLLang

	strCookie = Split(document.cookie,"&")

	for i = 0 to UBound(strCookie)
		if  instr(strCookie(i),"gURLLangUserID") <> 0  then	
			strURLLangUserID = mid(strCookie(i), (instr(1,strCookie(i),"=") +1))		
       		Exit For   
		end if 
	next
	
	strURLLangUserID = Split(strURLLangUserID, ";")
	URLLang = strURLLangUserID(0)
	HexCookie= Split(URLLang,"%")		
	
	for i = 1 to UBound(HexCookie)
		HexToString = "&H" & mid(HexCookie(i),1,2) 
		URLLangUsrID = URLLangUsrID & chr(CLng(HexToString)) & mid(HexCookie(i),3,len(HexCookie(i)))		
	next
	URLLangUsrID=replace(URLLangUsrID,"+"," ")
	GetURLLangUserID =  URLLangUsrID
End Function

'========================================================================================
' Function Name : GetURLLANG
' Function Desc : 현재 URL과 언어 코드 알아오기 
'========================================================================================
Function GetURLLANG()
	dim gURLLangPath
	Dim strLoc, iPos , iLoc, strPath, sepCount
	strLoc = window.location.href
		iLoc = inStr(1, strLoc, "?")
        If iLoc > 0 Then
			strLoc = Left(strLoc, iLoc - 1)
        End If
    
	iLoc = 1: iPos = 0 
	sepCount = 0
	
	Do Until sepCount > 4						
		iLoc = inStr(iPos+1, strLoc, "/")
		If iLoc <> 0 Then 
			iPos = iLoc
			sepCount= sepCount + 1
		end if
	Loop		
	gURLLangPath = Left(strLoc, iPos-1)
	GetURLLANG = gURLLangPath

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
Sub MainQuery()
    
    Call RestoreToolBar()
    Call DisableToolBar(TBC_QUERY)
    
    If FncQuery() = False then
	   Call RestoreToolBar()
	End If    
    
End Sub			

'========================================================================================
Sub MainNew()

    Call RestoreToolBar()
    Call DisableToolBar(TBC_NEW)
    
    Call FncNew()
    Call RestoreToolBar()
	
End Sub			

'========================================================================================
Sub MainDelete()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_DELETE)
    
    If FncDelete() = False then
	   Call RestoreToolBar()
	End If    

End Sub			

'========================================================================================
Sub MainSave()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_SAVE)
    
    If FncSave() = False then
	   Call RestoreToolBar()
	End If    
	
End Sub			

'========================================================================================
Sub MainInsertRow()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_INSERTROW)
    
	Call FncInsertRow()
	
    Call RestoreToolBar()
    
End Sub			

'========================================================================================
Sub MainDeleteRow()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_DELETEROW)
    
	Call FncDeleteRow()
	
    Call RestoreToolBar()
End Sub			

'========================================================================================
Sub MainPrev()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_PREV)
    
    If FncPrev() = False then
	   Call RestoreToolBar()
	End If    
	
End Sub			

'========================================================================================
Sub MainNext()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_NEXT)
    
    If FncNext() = False then
	   Call RestoreToolBar()
	End If    

End Sub			

'========================================================================================
Sub MainCancel()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_CANCEL)
    Call FncCancel()
    Call RestoreToolBar()
End Sub			

'========================================================================================
Sub MainCopy()
    Call RestoreToolBar()
    Call DisableToolBar(TBC_COPYRECORD)
    Call FncCopy()
    Call RestoreToolBar()
End Sub			

'========================================================================================
Sub MainExcel()
    Call FncExcel()

End Sub			

'========================================================================================
Sub MainPrint()
    Call FncPrint()
End Sub			

'========================================================================================
Sub MainFind()
    Call FncFind()
End Sub			

'======================================================================================================
Function MSIEVer()
   Dim tmpStr
   
   MSIEVer = "" 
   
   tmpStr = window.navigator.appVersion
   
   tmpStr = Split(tmpStr,";")    
   
   Select Case UCase(Trim(tmpStr(1)))
     Case "MSIE 5.5"
           MSIEVer = "5.5"
     Case "MSIE 6.0B"
           MSIEVer = "6.0B"
     Case "MSIE 6.0"
           MSIEVer = "6.0"
     Case "MSIE 7.0B"
           MSIEVer = "7.0B"
     Case "MSIE 7.0"
           MSIEVer = "7.0"
     Case "MSIE 8.0"
           MSIEVer = "8.0"
     Case "MSIE 9.0"
           MSIEVer = "9.0"
   End Select 
    
End Function

'======================================================================================================
Function PgmJump(strPgmID)	

    Call Parent.DBGo(strPgmID,False)
	'strVal = "../../ComAsp/Go.asp" & "?txtGo=" & strPgmID
	'Call RunMyBizASP(document, strVal)
End Function

'========================================================================================
Function OrProc(val1, val2)
	Dim i
	
	OrProc = ""
	
	val2 = Left(val2 & String(Len(val1), "0"), Len(val1))

	For i = 1 To Len(val1)
		If Mid(val1, i, 1) = "1" Or Mid(val2, i, 1) = "1" Then
			OrProc = OrProc & "1"
		Else
			OrProc = OrProc & "0"
		End If
	Next

End Function

'========================================================================================
' Function Name : RunMyBizASP
' Function Desc : 비지니스 로직 ASP에 Get 방식으로 실행시킨다.
'========================================================================================
Sub RunMyBizASP(objIFrame, strURL)
	Call BtnDisabled(True)
	objIFrame.location.href = GetUserPath & ValueEscape(strURL)
End Sub

'========================================================================================
Sub SetCOMProperty()
    
    On Error Resume Next    
 
    If gRdsUse = "T" Then    
       Call ggoOper.SetEnvData(gServerIP,gLogoName,gEnvInf,gRdsUse,gFontSize,gFontName)
    Else
       Call ggoOper.SetEnvData(GetURLLANG() & "/",gLogoName,escape(gEnvInf),gRdsUse,gFontSize,gFontName)
    End If

    Call ggoOper.SetNumberData(13,11,11,9)

    Call ggoSpread.SetSpreadFlagData("입력","수정","삭제")
    Call ggoSpread.SetEnvData(gLang,"6","Y",True,gRowSep,gColSep)
    Call ggoSpread.SetNumberData(13,11,11,9)
    Call ggoSpread.SetSpreadColor(UC_PROTECTED,UC_REQUIRED,RGB(209,232,249),"SYSTEM",&H333333,&HB7B7B7)
'    Call ggoSpread.SetSpreadColor(UC_PROTECTED,UC_REQUIRED,RGB(209,232,249),"SYSTEM")
    Call ggoSpread.SetXMLFileNameInf(gCompany,gDBServer,gDatabase,gUsrID)
    Call ggoSpread.SetSuperUser("unierp")
End Sub

'========================================================================================
Function SetToolBar(pVal)
	Dim strAuth
	Dim authVal
	
	strAuth = ""
	authVal = ""
	
	Err.Clear

	'gStrRequestMenuID : uni2kCM.inc file에서 정의된 프로그램을 요청한 메뉴의 ID
	If UCase(Left(gStrRequestMenuID, 1)) <> "Z" Then
		Call parent.uni2kMenu.Restore("BizMenu")
	Else
		Call parent.uni2kMenu.Restore("System")
	End If
	strAuth = parent.uni2kMenu.MenuItemAuthority(gStrRequestMenuID)
		
	Err.Clear

	Select Case strAuth
		Case "N"                         ' None
			authVal = "10000000000000"
		Case "Q"                         ' Query
			authVal = "11000000110001"
		Case "E"                         ' Excel/Print
			authVal = "11000000110111" 
		Case "A"                         ' All   
			authVal = "11111111111111"
	End Select

	If authVal = "" Then
		authVal = "10000000000000"
	End If
	
	authVal = AndProc(authVal, pVal)

	Call parent.SetToolbar(authVal)
	gToolBarBit = authVal
	
    Set gActiveElement = document.activeElement 
    
End Function

'========================================================================================
Sub RestoreResponseCookie(pMyBizASP,pLevel)
    Dim strVal
    
    On Error Resume Next
    Err.Clear

    Select Case pLevel
       Case 1
            strVal = "./"
       Case 2
            strVal = "../"
       Case 3
            strVal = "../../"
       Case 4
            strVal = "../../../"
    End Select    
       
    strVal = strVal & "PostSessionTrans.Asp?RedirectYN=" & "N"    

	Call RunMyBizASP(pMyBizASP, strVal)										'☜: 비지니스 ASP 를 가동	
	
End Sub	

'======================================================================================================
Sub CancelRestoreToolBar()
    gActionStatus = ""
End Sub	

'======================================================================================================
Sub RestoreToolBar()
    On Error Resume Next
    
    If gActionStatus = "" Then
       Exit Sub
    End If
    
    Select Case gActionStatus
       Case TBC_QUERY      : gToolBarBit = OrProc(gToolBarBit, "0100000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_NEW        : gToolBarBit = OrProc(gToolBarBit, "0010000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_DELETE     : gToolBarBit = OrProc(gToolBarBit, "0001000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_SAVE       : gToolBarBit = OrProc(gToolBarBit, "0000100000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_INSERTROW  : gToolBarBit = OrProc(gToolBarBit, "0000010000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_DELETEROW  : gToolBarBit = OrProc(gToolBarBit, "0000001000000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_CANCEL     : gToolBarBit = OrProc(gToolBarBit, "0000000100000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_PREV       : gToolBarBit = OrProc(gToolBarBit, "0000000010000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_NEXT       : gToolBarBit = OrProc(gToolBarBit, "0000000001000000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_COPYRECORD : gToolBarBit = OrProc(gToolBarBit, "0000000000100000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_EXPORT     : gToolBarBit = OrProc(gToolBarBit, "0000000000010000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_PRINT      : gToolBarBit = OrProc(gToolBarBit, "0000000000001000") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_FIND       : gToolBarBit = OrProc(gToolBarBit, "0000000000000100") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_HELP       : gToolBarBit = OrProc(gToolBarBit, "0000000000000010") : Call parent.SetToolbar(gToolBarBit)
       Case TBC_EXIT       : gToolBarBit = OrProc(gToolBarBit, "0000000000000001") : Call parent.SetToolbar(gToolBarBit)
    End Select 
    gActionStatus = ""
   
End Sub

'========================================================================================
Function UNIMsgBox(pVal, pType, pTitle)
	MsgBox pVal, pType, pTitle
End Function

'========================================================================================
Sub Window_onLoad()
    Dim iDx
    Dim ioriginal

 '  ioriginal = SetLocale(4103)
    
	Call GetGlobalVar
    Call SetCOMProperty                                                      '⊙: Load Common DLL
    
    gFocusSkip = False
    
    gMouseClickStatus = "N"

    gTabMaxCnt = 0
    gIsTab = "N"
    gPageNo =  1
    
    Call AdjustStyleSheet(Document)
    
    Call Form_Load()
    
    If Trim(gStrRequestUpperMenuID) <> "" Then
       top.document.title = gLogo & " - " & "[" & gStrRequestUpperMenuID & "][" & gStrRequestMenuID & "][" & document.title & "]"  
    Else
       top.document.title = gLogo & " - " & "[" &  document.title & "]"  
    End If

    window.status      = ""

    Set gActiveElement = document.activeElement 
    gLookUpEnable      = True    
    
    '------------------------------------------------------------------------------------
    ' Write current program id in cookie
    '------------------------------------------------------------------------------------
    iDx = Instr(UCase(document.location.href),"MODULE")

    If iDx > 0 And Trim(gStrRequestMenuID) > ""  Then   ' This means that if current program id is not popup    
       Document.Cookie = "gActivePgmID" & "=" & Mid(document.location.href,iDx ) & "; path=" & "/"    
    End If       
    
End Sub

'========================================================================================
Sub Window_onUnLoad()
	On Error Resume Next
	Dim Cancel, UnloadMode
	
	Call Form_QueryUnLoad( Cancel , UnloadMode)
	
 	Set gActiveElement = Nothing
End Sub

'========================================================================================
Function ExternalWrite(strData)
	Document.Write strData
End Function
