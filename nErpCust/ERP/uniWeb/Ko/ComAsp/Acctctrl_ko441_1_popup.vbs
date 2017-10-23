'========================================================================================
' Amount,Cost,Quantity,Exchange Rate Decimal Place Variable Definition
'========================================================================================
Class TB19029
     Dim DecPoint						'Decimal point
     Dim RndPolicy                      'Round Policy
     Dim RndUnit                        'Round Unit     
End Class

Dim ggAmtOfMoney						' Amount of Money
Set ggAmtOfMoney = New TB19029

Dim ggQty							' Quantity    
Set ggQty = New TB19029

Dim ggUnitCost						' Unit Cost
Set ggUnitCost = New TB19029

Dim ggExchRate						' Exchange Rate 
Set ggExchRate = New TB19029

DIm ggStrIntegeralPart                  ' Variable that contains value of  Integer Parts Places
DIm ggStrDeciPointPart                  ' Variable that contains value of  Decimal Parts Places

DIm ggStrMinPart                        ' Variable that contains minimum value
DIm ggStrMaxPart                        ' Variable that contains maximum value
Dim gURLPath 
Dim gIMGPath

Const UCN_REQUIRED      = "required"   'Required  field
Const UCN_PROTECTED     = "protected"  'Protected field
Const UCN_DEFAULT       = "normal"     'Optional  field

Const UC_REQUIRED_BAK   = &H99F7FF     'Color representing that Space should be Essentially input
Const UC_REQUIRED       = &HB4FFFF     'Color representing that Space should be Essentially input
Const UC_PROTECTED      = &Hdddddd     'Color representing that Space can not be input
Const UC_DEFAULT        = &HFFFFFF     'Color representing that Space can optionally be input


Dim gADODBConnString
Dim gColSep
Dim gComDateType
Dim gComNum1000
Dim gComNumDec
Dim gConnectionString
Dim gDateFormat
Dim gDateFormatYYYYMM
Dim gDsnNo
Dim gRowSep
Dim gLang
Dim gSeverity
Dim gUsrId
Dim gDBLoginPwd
Dim gDBServerIP
Dim gDBServerNm
Dim gClientIp
Dim gClientNm

Dim gCompany
Dim gDBServer
Dim gDatabase

Dim gRdsUse
Dim gEnvInf

'========================================================================================
'
'                                        PART - I
'
'========================================================================================

'========================================================================================
Sub Window_onLoad()
    Dim iDx
    Dim ioriginal

 	Call GetGlobalVar
    Call SetCOMProperty                                                      '⊙: Load Common DLL
    
    Call Form_Load()
    
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
	xmlDoc.async = False 
	
	xmlDoc.LoadXML(objConn("GlobalXMLData2"))
	
	NodeNm = "LoadBasisGlobalInf"
	
    gADODBConnString	 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gADODBConnString").text	
    gComDateType		 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComDateType").text	
    gComNum1000			 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNum1000").text
    gComNumDec           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNumDec").text	        
    gConnectionString    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gConnectionString").text
    gDateFormat          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormat").text
    gDateFormatYYYYMM	 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormatYYYYMM").text
    gDsnNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDsnNo").text	
    gLang                = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLang").text
    gSeverity            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSeverity").text
    gUsrId               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrId").text

    gCompany             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompany").text    '2005-04-26
    gDBServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServer").text   '2005-04-26
    gDatabase            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDatabase").text   '2005-04-26

    gColSep              = Chr(11)                    
    gRowSep			     = Chr(12)
    
    NodeNm ="GetGlobalInf"    
    
    gDBLoginPwd          = objConn("DBLoginPwd")
    gDBServerIP          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerIP").text		
    gDBServerNm          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerNm").text	    
    gClientIp            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientIp").text
    gClientNm            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNm").text

    gEnvInf =           gADODBConnString  & Chr(12)   '0
    gEnvInf = gEnvInf & gLang			  & Chr(12)   '1
    gEnvInf = gEnvInf & "CommonPopup"     & Chr(12)   '2
    gEnvInf = gEnvInf & gUsrId			  & Chr(12)   '3    
    gEnvInf = gEnvInf & gClientNm         &	Chr(12)   '4
    gEnvInf = gEnvInf & gClientIp         & Chr(12)   '5    
    gEnvInf = gEnvInf & gUsrId            & Chr(12)   '6
    gEnvInf = gEnvInf & gSeverity         & Chr(12)   '7

    gRdsUse = "F"
    
    set xmlDoc  = nothing
    set objConn = nothing
    
End Sub

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
Function CheckOCXQuery()
   
	CheckOCXQuery = False
	
    If MSIEVer() = "5.5" Then
       CheckOCXQuery = True
       Exit Function
    End If
	
End Function

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
' Function Name : ExecMyBizASP
' Function Desc : 비지니스 로직 ASP에 Post 방식으로 실행시킨다.
'========================================================================================
Sub ExecMyBizASP(objForm, strAction)
'   Call BtnDisabled(True)
	objForm.action = GetUserPath & strAction
'	Call elementEnabled(True)
	objForm.submit
'	Call elementEnabled(False)
End Sub

'========================================================================================
' Function Name : RunMyBizASP
' Function Desc : 비지니스 로직 ASP에 Get 방식으로 실행시킨다.
'========================================================================================
Sub RunMyBizASP(objIFrame, strURL)
'	Call BtnDisabled(True)
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
  ' Call ggoSpread.SetSpreadColor(UC_PROTECTED,UC_REQUIRED,RGB(209,232,249),"SYSTEM",&H333333,&HB7B7B7)
    Call ggoSpread.SetSpreadColor(UC_PROTECTED,UC_REQUIRED,RGB(209,232,249),"SYSTEM",&H333333,&HB7B7B7,&HF1EFED)  '2005-04-09
    Call ggoSpread.SetXMLFileNameInf(gCompany,gDBServer,gDatabase,gUsrID)
    Call ggoSpread.SetSuperUser("unierp")
End Sub


'======================================================================================================
Function VisibleRowCnt(pDoc,ByVal pStartRow)
    Dim i,j
    Dim pStartCol
    
    On Error Resume Next
    
    VisibleRowCnt = 0
    
    If pStartRow < 0 Then 
       Exit Function
    End If

    If pStartRow = 0 Then
       pStartRow = 1
       VisibleRowCnt =  50 
       Exit Function
    End If
    
    pStartCol = 1
    
    For i = 1 To  pDoc.MaxCols
        pDoc.Col = i
        pDoc.Row = pStartRow
        If pDoc.ColHidden <> True Then
           pStartCol = i
           Exit For
        End If
    Next    
    
    For i = pStartCol To  pDoc.MaxCols                                    ' Left to Right
        If pDoc.IsVisible(i, pStartRow, True) = True Then
           pStartCol = i
           Exit For
        End If
    Next    

    If pDoc.IsVisible(pStartCol, pStartRow, True) = False Then           ' Top to Bottom
       For i = pStartRow To pDoc.MaxRows
           If pDoc.IsVisible(pStartCol, i, True) = True Then
              pStartRow = i
              Exit For
           End If
       Next
    End If

    For i = pStartRow To pDoc.MaxRows                                  ' Count visible row
       If pDoc.IsVisible(pStartCol, i, True) = False Then
          Exit For
       End If
       j = i
    Next
    
    VisibleRowCnt = j - pStartRow + 1
 
End Function


'======================================================================================================
' Function Name : ValueEscape
' Function Desc : GET 방식으로 넘어가는 문자열에서 name=value 에서 value를 escape 시킨다 
'======================================================================================================
Function ValueEscape(strURL)
	Dim szTarget, szAmp
	Dim szValz, szTemp
	Dim arrToken()
	Dim s, e, i, nCnt, s1, e1
    Dim ii,sp
    Dim iTmp, iTmpA, iTmpB, iTmpC, iTmpArr
    Dim iReplaceChar

    iReplaceChar = "**4***4**"
    
	s = Instr(1, strURL, "?")
	If s = 0 Then
		ValueEscape = strURL
		Exit Function
	End If

	szTarget = Left(strURL, s)
	szValz = Mid(strURL, s+1, Len(strURL) - s + 1)
	
    iTmpArr = Split(szValz, "=")
  
    For ii = 0 To UBound(iTmpArr)
        iTmpArr(ii) = StrReverse(iTmpArr(ii))
    Next
  
    For ii = 0 To UBound(iTmpArr) - 1
        iTmp = InStr(iTmpArr(ii), "&")
        If iTmp > 0 Then
           iTmpA = Mid(iTmpArr(ii), 1, iTmp - 1)
           iTmpB = Mid(iTmpArr(ii), iTmp + 1)
           iTmpArr(ii) = iTmpA & "&" & Replace(iTmpB, "&", iReplaceChar)
        End If
    Next
  
    iTmpArr(ii) = Replace(iTmpArr(ii), "&", iReplaceChar)
  
    szValz = StrReverse(iTmpArr(0))
  
    For ii = 1 To UBound(iTmpArr)
        szValz = szValz & "=" & StrReverse(iTmpArr(ii))
    Next

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

        arrToken(1, i) = escape(arrToken(1, i))
		
		arrToken(1, i) = Replace(arrToken(1, i), "+", "%2B")
		arrToken(1, i) = Replace(arrToken(1, i), "/", "%2F")

		ValueEscape = ValueEscape + szAmp + arrToken(0, i) + "=" + arrToken(1, i)
		
    Next

    ValueEscape = Replace(ValueEscape, iReplaceChar, "%26")

End Function

'======================================================================================================
' Function Name : PreEscape
' Function Desc : Form의 컨트롤에서 받는 값을 escape 시킨다 
'======================================================================================================
Function PreEscape(ByVal strVal)
	PreEscape = Escape(strVal)
End Function


'========================================================================================
' Sub Name : LayerShowHide(Status)
' Sub Desc : 마우스 포인터용 Layer의 Visibility 설정 
'========================================================================================
Function LayerShowHide(ByVal Status)
	Dim LayerN

	On Error Resume Next
	
	LayerShowHide = False
	
	If Status = 0 Then 
		Status = "hidden"
	Else
		Status = "visible"
	End If

	Set LayerN = window.document.all("MousePT").style

	If Err.Number = 0 Then 
	    If LayerN.visibility = Status And Status = "visible" Then
'	       Exit Function
	    End If
	
		LayerN.visibility = Status
	Else
		Err.Clear				    			
	End if		

	LayerShowHide = True

End Function 

'======================================================================================================
Function DisplayMsgBoxA(ByVal pMsgId)
    Dim iCount
    Dim iRet
    Dim iRet1
    Dim iRet2
    Dim iRet3

    If FetchBMessage(Cstr(pMsgId),iRet ) = True Then
    Else
       Exit Function         
    End If
  
    iRet2 = Split(iRet,Chr(12))
    
    If iRet2(0) = "X" Then
       iRet2(2) = "1"             'Default value set
    End If

    Select Case iRet2(2)
            Case "1"   ' Information
                   DisplayMsgBox = MsgBox(iRet2(1), vbInformation, gLogoName & "-[Information]")
            Case "2"   ' Warning
                   DisplayMsgBox = MsgBox(iRet2(1), vbExclamation, gLogoName & "-[Warning]")
            Case "3"   ' Error
                   DisplayMsgBox = MsgBox(iRet2(1), vbExclamation, gLogoName & "-[Error]")
            Case "4"   ' Fatal   
                   DisplayMsgBox = MsgBox(iRet2(1), vbCritical   , gLogoName & "-[Fatal]")
    End Select

End Function


'======================================================================================================
'Function Name   : FetchBMessage
'Function Desc   : This function query message text according to the message code
'Return   value  : return status code + message text + message severity
'======================================================================================================
Function FetchBMessage(pCode,prData)
    Dim iXmlHttp
    Dim iSendStr
    Dim pRDSCom
    Dim iRetByte

    On Error Resume Next

    FetchBMessage =  False
      
        Set iXmlHttp = CreateObject("Msxml2.XMLHTTP")		

        iXmlHttp.open "POST", GetComaspFolderPath & "RequestGetMSG.asp", False     
        iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
       
        iSendStr = escape(gEnvInf)
        iSendStr = Replace(iSendStr, "+", "%2B")
        iSendStr = Replace(iSendStr, "/", "%2F")
    
        iXmlHttp.send "iMode=GET&LangCD=" & glang & "&MsgCd=" & pCode & "&EnvInf=" & iSendStr

        If gCharSet = "D" Then 'U : unicode, D:DBCS
           prData   = ConnectorControl.CStrConv(iXmlHttp.responseBody)
        Else
           prData   = iXmlHttp.responseText
        End If   

        Set iXmlHttp = Nothing           
    
    If Err.number <> 0 Then
        Exit Function
    End If
    
    FetchBMessage = True
    
End Function

'========================================================================================
' Function Name : Document_onKeyDown
' Function Desc : hand all event of key down
'========================================================================================
Function Document_onKeyDown()
	Dim objEl, KeyCode, iLoc
	Dim boolMinus, boolDot
	
	On Error Resume Next
	
	Document_onKeyDown = True
	Set objEl = window.event.srcElement
	KeyCode   = window.event.keycode
	
	Select Case KeyCode	
		Case 13		' Enter Key: Used as Query in Condition

				If Left(objEl.getAttribute("tag"),1) = "1" Then
								
                      If UCase(objEl.tagName) = "OBJECT" And CheckOCXQuery = False Then
				      Else
                         Call FncQuery()
                      End If 				
                      Exit Function
				End If
				
		Case 27   'ESC
		       Self.Close
               Exit Function
	End Select
	
End Function

'========================================================================================
Function ExternalWrite(strData)
	Document.Write strData
End Function
