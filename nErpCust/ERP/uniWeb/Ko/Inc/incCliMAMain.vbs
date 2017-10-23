Dim gURLPath 
Dim gIMGPath
'========================================================================================
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
       Case Parent.TBC_QUERY      : iBit = "1011111111111111"
       Case Parent.TBC_NEW        : iBit = "1101111111111111"
       Case Parent.TBC_DELETE     : iBit = "1110111111111111"
       Case Parent.TBC_SAVE       : iBit = "1111011111111111"
       Case Parent.TBC_INSERTROW  : iBit = "1111101111111111"
       Case Parent.TBC_DELETEROW  : iBit = "1111110111111111"
       Case Parent.TBC_CANCEL     : iBit = "1111111011111111"
       Case Parent.TBC_PREV       : iBit = "1111111101111111"
       Case Parent.TBC_NEXT       : iBit = "1111111110111111"
       Case Parent.TBC_COPYRECORD : iBit = "1111111111011111"
       Case Parent.TBC_EXPORT     : iBit = "1111111111101111"
       Case Parent.TBC_PRINT      : iBit = "1111111111110111"
       Case Parent.TBC_FIND       : iBit = "1111111111111011"
       Case Parent.TBC_HELP       : iBit = "1111111111111101"
       Case Parent.TBC_EXIT       : iBit = "1111111111111110"
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

'========================================================================================
Sub GetGlobalVar()
    gEnvInf =           parent.gADODBConnString & Chr(12)   '0
    gEnvInf = gEnvInf & parent.gLang		    & Chr(12)   '1
    gEnvInf = gEnvInf & gStrRequestMenuID       & Chr(12)   '2
    gEnvInf = gEnvInf & parent.gUsrId           & Chr(12)   '3    
    gEnvInf = gEnvInf & parent.gClientNm        & Chr(12)   '4
    gEnvInf = gEnvInf & parent.gClientIp        & Chr(12)   '5    
    gEnvInf = gEnvInf & parent.gUsrId           & Chr(12)   '6
    gEnvInf = gEnvInf & parent.gSeverity        & Chr(12)   '7
    
    gADODBConnString = parent.gADODBConnString       
    gDsnNo           = parent.gDsnNo
    gDateFormat      = parent.gDateFormat
    gLang            = parent.gLang

    gServerIP        = parent.gServerIP
    gLogoName        = parent.gLogoName 
    gLogo            = parent.gLogo
    gRdsUse          = parent.gRdsUse   
    gCharSet         = parent.gCharSet
    gCharSQLSet      = parent.gCharSQLSet
   
End Sub

'========================================================================================
Function GetCustGlobalVar()
    Dim objConn
    
    Set objConn = CreateObject("uniConnector.cGlobal")
    
    If Err.Number = 0 Then
       GetCustGlobalVar = objConn("CUSTXMLSTR")
    End If

End Function
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

'========================================================================================
Sub MainQuery()
    
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_QUERY)
    
    If FncQuery() = False then
	   Call RestoreToolBar()
	End If    
    
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainNew()

    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_NEW)
    
    Call FncNew()
    Call RestoreToolBar()
	
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainDelete()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_DELETE)
    
    If FncDelete() = False then
	   Call RestoreToolBar()
	End If    

    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainSave()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_SAVE)
    
    If FncSave() = False then
	   Call RestoreToolBar()
	End If    
	
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainInsertRow()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_INSERTROW)
    
	Call FncInsertRow("1")
	
    Call RestoreToolBar()
    
    Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainDeleteRow()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_DELETEROW)
    
	Call FncDeleteRow()
	
    Call RestoreToolBar()
    Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainPrev()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_PREV)
    
    If FncPrev() = False then
	   Call RestoreToolBar()
	End If    
	
    Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainNext()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_NEXT)
    
    If FncNext() = False then
	   Call RestoreToolBar()
	End If    
    Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainCancel()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_CANCEL)
    Call FncCancel()
    Call RestoreToolBar()
    Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainCopy()
    Call RestoreToolBar()
    Call DisableToolBar(Parent.TBC_COPYRECORD)
    Call FncCopy()
    Call RestoreToolBar()
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainExcel()
    Call FncExcel()
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainPrint()
    Call FncPrint()
    'Set gActiveElement = document.activeElement
End Sub			

'========================================================================================
Sub MainFind()
    Call FncFind()
    'Set gActiveElement = document.activeElement
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
       Call ggoOper.SetEnvData(Parent.gServerIP,Parent.gLogoName,gEnvInf,gRdsUse,Parent.gFontSize,Parent.gFontName,"T")
    Else
       Call ggoOper.SetEnvData(GetRootFolderByLang(),Parent.gLogoName,escape(gEnvInf),gRdsUse,Parent.gFontSize,Parent.gFontName,"T")
    End If
    
    Call ggoOper.SetNumberData(13,11,11,9)

    Call ggoSpread.SetSpreadFlagData("입력","수정","삭제")
    Call ggoSpread.SetEnvData(Parent.gLang,"6","Y",True,Parent.gRowSep,Parent.gColSep)
    Call ggoSpread.SetNumberData(13,11,11,9)
'   Call ggoSpread.SetSpreadColor(Parent.UC_PROTECTED,Parent.UC_REQUIRED,RGB(1,232,249),"SYSTEM",&H333333,&HB7B7B7)   
    Call ggoSpread.SetSpreadColor(Parent.UC_PROTECTED,Parent.UC_REQUIRED,RGB(1,232,249),"SYSTEM",&H333333,&HB7B7B7,&HF1EFED)  '2005-04-09
    Call ggoSpread.SetXMLFileNameInf(Parent.gCompany,Parent.gDBServer,Parent.gDatabase,Parent.gUsrID)
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
' Function Name : RestoreResponseCookie()
' Function Desc : Restore Response Cookie
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
       Case Parent.TBC_QUERY      : gToolBarBit = OrProc(gToolBarBit, "0100000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_NEW        : gToolBarBit = OrProc(gToolBarBit, "0010000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_DELETE     : gToolBarBit = OrProc(gToolBarBit, "0001000000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_SAVE       : gToolBarBit = OrProc(gToolBarBit, "0000100000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_INSERTROW  : gToolBarBit = OrProc(gToolBarBit, "0000010000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_DELETEROW  : gToolBarBit = OrProc(gToolBarBit, "0000001000000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_CANCEL     : gToolBarBit = OrProc(gToolBarBit, "0000000100000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_PREV       : gToolBarBit = OrProc(gToolBarBit, "0000000010000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_NEXT       : gToolBarBit = OrProc(gToolBarBit, "0000000001000000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_COPYRECORD : gToolBarBit = OrProc(gToolBarBit, "0000000000100000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_EXPORT     : gToolBarBit = OrProc(gToolBarBit, "0000000000010000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_PRINT      : gToolBarBit = OrProc(gToolBarBit, "0000000000001000") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_FIND       : gToolBarBit = OrProc(gToolBarBit, "0000000000000100") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_HELP       : gToolBarBit = OrProc(gToolBarBit, "0000000000000010") : Call parent.SetToolbar(gToolBarBit)
       Case Parent.TBC_EXIT       : gToolBarBit = OrProc(gToolBarBit, "0000000000000001") : Call parent.SetToolbar(gToolBarBit)
    End Select 
    gActionStatus = ""
   
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
    Dim ioriginal
    
    On Error Resume Next
    
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
       top.document.title = parent.gLogo & " - " & "[" & gStrRequestUpperMenuID & "][" & gStrRequestMenuID & "][" & document.title & "]"  
    Else
       top.document.title = parent.gLogo & " - " & "[" &  document.title & "]"  
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

