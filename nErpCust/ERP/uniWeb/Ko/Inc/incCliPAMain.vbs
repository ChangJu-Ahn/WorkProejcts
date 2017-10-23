Dim gURLPath 
Dim gIMGPath

'========================================================================================
Function CheckOCXQuery()
   
	CheckOCXQuery = False
	
    If MSIEVer() = "5.5" Then
       CheckOCXQuery = True
       Exit Function
    End If
	
End Function

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

    gEnvInf =           PopupParent.gADODBConnString & Chr(12)   '0
    gEnvInf = gEnvInf & PopupParent.gLang	         & Chr(12)   '1
    gEnvInf = gEnvInf & "Popup"                      & Chr(12)   '2
    gEnvInf = gEnvInf & PopupParent.gUsrId           & Chr(12)   '3    
    gEnvInf = gEnvInf & PopupParent.gClientNm        & Chr(12)   '4
    gEnvInf = gEnvInf & PopupParent.gClientIp        & Chr(12)   '5    
    gEnvInf = gEnvInf & PopupParent.gUsrId           & Chr(12)   '6
    gEnvInf = gEnvInf & PopupParent.gSeverity        & Chr(12)   '7

    gADODBConnString = PopupParent.gADODBConnString       
    gDsnNo           = PopupParent.gDsnNo
    gDateFormat      = PopupParent.gDateFormat
    gLang            = PopupParent.gLang
            
    gServerIP        = PopupParent.gServerIP
    gLogoName        = PopupParent.gLogoName 
    gLogo            = PopupParent.gLogo
    gRdsUse          = PopupParent.gRdsUse  

    gCharSet         = PopupParent.gCharSet
    gCharSQLSet      = PopupParent.gCharSQLSet
     
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
Function GetRootFolderByLang()
   Dim iStrTemp
   Dim iPath   
   Dim i
   
   iStrTemp = Document.Location.href
   iStrTemp = Split(iStrTemp,"/")
   
   For i = 0 To 4
      iPath = iPath & iStrTemp(i) & "/"
   Next
   GetRootFolderByLang = iPath
End Function

'======================================================================================================
' Name : MSIEVer
' Desc : Get Current Internet Explorer version
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
   End Select 
    
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
        Call ggoOper.SetEnvData(PopupParent.gServerIP,PopupParent.gLogoName,gEnvInf,gRdsUse,PopupParent.gFontSize,PopupParent.gFontName)
    Else
        Call ggoOper.SetEnvData(GetRootFolderByLang(),PopupParent.gLogoName,escape(gEnvInf),gRdsUse,PopupParent.gFontSize,PopupParent.gFontName)
    End If
    Call ggoOper.SetNumberData(13,11,11,9)

    Call ggoSpread.SetSpreadFlagData("입력","수정","삭제")
    Call ggoSpread.SetEnvData(PopupParent.gLang,"6","Y",True,PopupParent.gRowSep,PopupParent.gColSep)
    Call ggoSpread.SetNumberData(13,11,11,9)
  '  Call ggoSpread.SetSpreadColor(PopupParent.UC_PROTECTED,PopupParent.UC_REQUIRED,RGB(209,232,249),"SYSTEM",&H333333,&HB7B7B7)
    Call ggoSpread.SetSpreadColor(PopupParent.UC_PROTECTED,PopupParent.UC_REQUIRED,RGB(1,232,249),"SYSTEM",&H333333,&HB7B7B7,&HF1EFED)  '2005-04-09
    
    Call ggoSpread.SetXMLFileNameInf(PopupParent.Company,PopupParent.gDBServer,PopupParent.gDatabase,PopupParent.gUsrID)
    Call ggoSpread.SetSuperUser("unierp")
End Sub

'========================================================================================
Function UNIMsgBox(pVal, pType, pTitle)
	MsgBox pVal, pType, pTitle
End Function

'========================================================================================
' Function Name : Window_onLoad
' Function Desc : 화면 처리 ASP가 클라이언트에 Load된 후 실행해야 될 로직 처리 
'========================================================================================
Sub Window_onLoad()
    Dim iDx
'    Dim iTimer
    Dim ioriginal

 '   ioriginal = SetLocale(4103)
    
	Call GetGlobalVar
    Call SetCOMProperty                                                      '⊙: Load Common DLL
    
    gFocusSkip = False
    
    gMouseClickStatus = "N"

    gTabMaxCnt = 0
    gIsTab = "N"
    gPageNo =  1
    
    Call AdjustStyleSheet(Document)
    
'    iTimer = Timer
    Call Form_Load()
'    Call MakeMainLog(Timer - iTimer)    

    If Trim(gStrRequestUpperMenuID) <> "" Then
       top.document.title = PopupParent.gLogo & " - " & "[" & gStrRequestUpperMenuID & "][" & gStrRequestMenuID & "][" & document.title & "]"  
    Else
       top.document.title = PopupParent.gLogo & " - " & "[" &  document.title & "]"  
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
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Window_onUnLoad()
	On Error Resume Next
	Dim Cancel, UnloadMode
	
	Call Form_QueryUnLoad( Cancel , UnloadMode)
	
 	Set gActiveElement = Nothing
End Sub

'========================================================================================
' Name :
'
'========================================================================================
Sub MakeMainLog(ByVal pTimer)

    Dim iXmlHttp
    Dim iSTR1,iSTR2

	On Error Resume Next

	Err.Clear
	
    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP")		
    iSTR1 = document.location.href
    iSTR1 = Split(iSTR1,"/")
    iSTR2 = Split(iSTR1(ubound(iSTR1)),".")

    iXmlHttp.open "GET", GetComaspFolderPath & "ComLogger.asp?ClientIp=" & popupparent.gClientIp & "&Timer=" & pTimer  & "&MenuID=" & iSTR2(0) & "&UsrID=" & popupparent.gUsrID  , False        

    iXmlHttp.send
    Set iXmlHttp = Nothing
End Sub 