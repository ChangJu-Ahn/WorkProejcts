
'========================================================================================
' Description   : Variables
'========================================================================================
Dim ObjName
Dim pType
Dim pEbId
Dim EBParent
'========================================================================================
' Function Name : GetParent()
' Description   : Get parent information
'========================================================================================
Sub GetParent

	If IsEmpty(PopupParent) then
		Set EBParent = Parent
	Else
		Set EBParent = PopupParent
	End If
	
End Sub
'========================================================================================
' Function Name : setPolicy()
' Description   : make condvar string (system info)
'========================================================================================
Function setPolicy

	Dim strVar
	Dim strDate
	Dim strMonth
	Dim AmtRndPolicy
	Dim QtyRndPolicy
	Dim UnitCostRndPolicy
	Dim ExchRateRndPolicy
	Dim ArrCur, ArrDecPoint, ArrRndPolicy, ArrDefaultDec, ArrTemp
	Dim i, iArrCount	
	
	StrDate = Replace(EBParent.gDateFormat,"Y", "y")    ' Y M D 를 대소문자로 구분 
	StrDate = Replace(StrDate,    "m", "M")
	StrDate = Replace(StrDate,    "D", "d")
		
	StrMonth = Replace(EBParent.gDateFormatYYYYMM,"Y", "y")    ' Y M 를 대소문자로 구분 
	StrMonth = Replace(StrMonth,         "m", "M")
	
	AmtRndPolicy = Replace(ggAmtofMoney.RndPolicy, "1", "+")  '금액 반올림정책 세팅 
	AmtRndPolicy = Replace(AmtRndPolicy, "2", "-")
	AmtRndPolicy = Replace(AmtRndPolicy, "3", "%20")
	
	QtyRndPolicy = Replace(ggQty.RndPolicy, "1", "+")          '수량 반올림정책 세팅 
	QtyRndPolicy = Replace(QtyRndPolicy, "2", "-")
	QtyRndPolicy = Replace(QtyRndPolicy, "3", "%20")
	
	UnitCostRndPolicy = Replace(ggUnitCost.RndPolicy, "1", "+")  '단가 반올림정책 세팅 
	UnitCostRndPolicy = Replace(UnitCostRndPolicy, "2", "-")
	UnitCostRndPolicy = Replace(UnitCostRndPolicy, "3", "%20")
	
	ExchRateRndPolicy = Replace(ggExchRate.RndPolicy, "1", "+")  '환율 반올림정책 세팅 
	ExchRateRndPolicy = Replace(ExchRateRndPolicy, "2", "-")
	ExchRateRndPolicy = Replace(ExchRateRndPolicy, "3", "%20")

    If IsArray(gBDataType) Then
		iArrCount = UBound(gBCurrency) - 1
		ReDim ArrCur(iArrCount)
    
		For i = 0 To iArrCount
		    ArrCur(i) = gBCurrency(i) & gBDataType(i)
		Next
	
		ArrCur = Join(ArrCur,chr(5))
		ArrDecPoint = Join(gBDecimals,chr(5))
		ArrRndPolicy = Join(gBRoundingPolicy,chr(5))
		ArrRndPolicy = Replace(ArrRndPolicy, "1", "+")  '환율 반올림정책 세팅 
		ArrRndPolicy = Replace(ArrRndPolicy, "2", "-")
		ArrRndPolicy = Replace(ArrRndPolicy, "3", "%20")
		ArrTemp = Split(ggStrDeciPointPart,chr(11))
		ArrDefaultDec = ArrTemp(10) & Chr(5) & ArrTemp(11) & Chr(5) & ArrTemp(12) & Chr(5) & ArrTemp(13)
	Else
	    ArrCur = ""
	    ArrDecPoint = ""
	    ArrRndPolicy = ""
	    ArrDefaultDec = "2" & Chr(5) & "4" & Chr(5) &  "4" & Chr(5) & "6" 
	End If
	
	strVar = strVar & "|AmtDecPoint|"       & ggAmtOfMoney.DecPoint
	strVar = strVar & "|QtyDecPoint|"       & ggQty.DecPoint
	strVar = strVar & "|UnitCostDecPoint|"  & ggUnitCost.DecPoint
	strVar = strVar & "|ExchRateDecPoint|"  & ggExchRate.DecPoint
	strVar = strVar & "|AmtRndPolicy|"      & AmtRndPolicy
	strVar = strVar & "|QtyRndPolicy|"      & QtyRndPolicy
	strVar = strVar & "|UnitCostRndPolicy|" & UnitCostRndPolicy
	strVar = strVar & "|ExchRateRndPolicy|" & ExchRateRndPolicy
	strVar = strVar & "|Num1000|"           & EBParent.gComNum1000
	strVar = strVar & "|DateFormat|"        & StrDate
	strVar = strVar & "|MonthFormat|"       & StrMonth 
	strVar = strVar & "|ArrCur|"            & ArrCur 
	strVar = strVar & "|ArrDecPoint|"       & ArrDecPoint
	strVar = strVar & "|ArrRndPolicy|"      & ArrRndPolicy
	strVar = strVar & "|ArrDefaultDec|"     & ArrDefaultDec
	strVar = strVar & "|gAlignOpt|"         & EBParent.gQMDPAlignOpt & "|"
	
    SetPolicy = strVar

End Function

'========================================================================================
' Function Name : GetPath(URL)
' Description   : Get path of 'lang'
'========================================================================================
Function GetPath(URL)

	Dim arrPath
	Dim PATH
	
	arrPath = split(URL,"/")   ' URL is http://servername/Company_db/lang....
	
	PATH = "HTTP://" & arrPath(2) & "/" & arrpath(3) & "/"  & arrPath(4) & "/"
	
	GetPath = PATH
	
End Function

'========================================================================================
' Function Name : GetEBVer()
' Description   : Get version of Reqube
'========================================================================================
Function GetEBVer()

	Dim EBVer, tmpVer

	Call GetParent
	
	EBVer = split(EBParent.gEbEnginePath5,"/")

	tmpVer = UCase(replace(EBVer(1), chr(11), ""))

	GetEBver = tmpVer

End Function

'========================================================================================
' Function Name : FncEBR5(EBRName, work, condvar, x, y)  -- work : "view", "print"
' Description   : Easybase web Preview for REQUBE 5.1 (EBR) 
'========================================================================================
Sub FncEBR5(EBRName, work, condvar, x, y, x1, y1)

	Call FncEBR5RC(EBRName, work, condvar, x, y, x1, y1, "EBR")
	
End Sub

'========================================================================================
' Function Name : FncEBR5(EBRName, work, condvar, x, y)  -- work : "view", "print"
' Description   : Easybase web Preview for REQUBE 5.1 (EBR) 
'========================================================================================
Sub FncEBR5RC(EBRName, work, condvar, x, y, x1, y1, EBRorEBC)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(6), arrField, arrHeader
	Dim pre
	Dim iASP

	pre = GetPath(window.location.href)
		
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath5
	strUrl = strUrl & work
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbUserPass   
	arrParam(2) = EBParent.gEbPkgRptPath5 & "/" & gLang & "/" & EBRorEBC & "/" & EBParent.gEbDbName & "/" & EBRName
	arrParam(3) = condvar & strPolicy
	arrParam(4) = strUrl
	arrParam(5) = work
	
	Select Case Trim(UCase(EBRorEBC))
	   Case "EBC" : iASP = "EBCpreview5.asp"
	   Case "EBR" : iASP = "EBRpreview5.asp"
	End Select   

	Call BtnDisabled(1)	

	window.showModalDialog  pre & "comasp/" & iASP & "?workFlag=" & work & "&dialogWidth=" & x & "&dialogHeight=" & y & "&EBWidth=" & x1 & "&EBHeight=" & y1 , Array(arrParam, arrField, arrHeader), _
	"dialogWidth="& x & "px; dialogHeight=" & y &"px; center: Yes; help: No; resizable: Yes; status: No; scroll:no;"

	Call BtnDisabled(0)	
	
End Sub


'========================================================================================
' Function Name : FncEBR5(EBRName, work, condvar, x, y)  -- work : "view", "print"
' Description   : Easybase web Preview for REQUBE 5.1 (EBR) 
'========================================================================================
Sub FncEBR5RC2(EBRName, work, condvar,pTarget,ByVal EBRorEBC)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(6), arrField, arrHeader
	Dim pre
	Dim iASP
	Dim uname,pw,filename
	
	Call GetParent

	pre = GetPath(window.location.href)
		
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath5
	strUrl = strUrl & work
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbUserPass   
	arrParam(2) = EBParent.gEbPkgRptPath5 & "/" & gLang & "/" & EBRorEBC & "/" & EBParent.gEbDbName & "/" & EBRName
	
	arrParam(3) = condvar & strPolicy
	arrParam(4) = strUrl
	arrParam(5) = work
	
	Select Case Trim(UCase(EBRorEBC))
	   Case "EBC" : iASP = "EBCpreview5.asp"
	   Case "EBR" : iASP = "EBRpreview5.asp"
	End Select   

	Call BtnDisabled(1)	
	
  	uname     = arrparam(0)
	pw        = arrparam(1)
	filename  = arrparam(2)
	condvar   = arrparam(3)
	strUrl    = arrparam(4)
	work      = arrParam(5)

	pTarget.pw.value      = pw
	pTarget.id.value     = uname
	pTarget.doc.value     = filename
	pTarget.runvar.value  = condvar
	
	pTarget.form.value   = "ACTIVEX"
	
	pTarget.action         = strUrl
	pTarget.submit


	
End Sub

'========================================================================================
Sub FncEBCPrint(objForm, EBRName, condvar)

    If GetEBVer = "REQUBE" Then
       Call FncEBR5RC(EBRName, "print", condvar, 280, 100, 280, 100,"EBC")
       Exit Sub
    End If
    
    Call FncPrint(objForm, EBRName, condvar)
    
End Sub

'========================================================================================
Sub FncEBRPrint(objForm, EBRName, condvar)

    If GetEBVer = "REQUBE" Then
       Call FncEBR5(EBRName, "print", condvar, 280, 100, 280, 100)
       Exit Sub
    End If
    
    Call FncPrint(objForm, EBRName, condvar)
    
End Sub

'========================================================================================
' Function Name : FncEBCPreview(EBCName, condvar)
' Description   : Easybase web Preview  (EBC) 
'========================================================================================

Sub FncEBCPreview(EBCName, condvar)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(6), arrField, arrHeader
	Dim x, y  'size
	Dim pre

	if GetEBVer = "REQUBE" then

	   Call AskEBDocumentName2(Split(EBCName,".")(0),x,y,x1,y1)
	   Call FncEBR5RC(EBCName, "view", condvar, x, y, x1, y1, "EBC")
	   Exit Sub	
	end if

	pre = GetPath(window.location.href)
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath
	strUrl = strUrl & "ExecuteWinCrossTab"
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbDbName
	arrParam(2) = EBParent.gEbPkgRptPath & "\" & gLang & "\EBC\" & EBCName
	arrParam(3) = condvar & strPolicy
	arrParam(4) = "-2"
	arrParam(5) = strUrl
	arrParam(6) = "1"
	
	x = 800
	y = 700
	Call BtnDisabled(1)	

	window.showModalDialog pre & "comasp/EBCpreview.asp?EBWidth=" & x & "&EBHeight=" & y , Array(arrParam, arrField, arrHeader), _
	"dialogWidth="& x & "px; dialogHeight=" & y &"px; center: Yes; help: No; resizable: Yes; status: No; scroll:no;"

	Call BtnDisabled(0)	
	
End Sub

'========================================================================================
' Function Name : FncEBRPreview(EBRName, condvar)
' Description   : Easybase web Preview  (EBR) 
'========================================================================================

Sub FncEBRPreview(EBRName, condvar)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(5), arrField, arrHeader
	Dim x, y                                                  'size
	Dim pre
	
	If GetEBVer = "REQUBE" then
	   Call AskEBDocumentName2(Split(EBRName,".")(0),x,y,x1,y1)
	   
	   Call FncEBR5(EBRName, "view", CondVar, x,y,x1,y1)
	   Exit Sub	
	End If

	pre = GetPath(window.location.href)
		
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath
	strUrl = strUrl & "ExecuteWinReport"
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbDbName
	arrParam(2) = EBParent.gEbPkgRptPath & "/" & gLang & "/EBR/" & EBRName
	arrParam(3) = condvar & strPolicy
	arrParam(4) = "-2"
	arrParam(5) = strUrl

	
	x = 800
	y = 700

	Call BtnDisabled(1)	
	
	window.showModalDialog  pre & "comasp/EBRpreview.asp?EBWidth=" & x & "&EBHeight=" & y , Array(arrParam, arrField, arrHeader), _
	"dialogWidth="& x & "px; dialogHeight=" & y &"px; center: Yes; help: No; resizable: Yes; status: No; scroll:no;"

	Call BtnDisabled(0)	
	
End Sub









'========================================================================================
Sub FncPrint(objForm, EBRName, condvar)

    Dim strUrl
    Dim strPolicy
    Dim arrParam, arrField, arrHeader
    
    strPolicy = setPolicy()
    
    strUrl = strUrl & EBParent.gServerIP
    strUrl = strUrl & EBParent.gEbEnginePath
    strUrl = strUrl & "ExecuteWinReportForPrint"
    
    objForm.uname.Value = EBParent.gEbUserName
    objForm.dbname.Value = EBParent.gEbDbName
    objForm.FileName.Value = EBParent.gEbPkgRptPath & "\" & gLang & "\EBR\" & EBRName
    objForm.condvar.Value = condvar & strPolicy
    objForm.Date.Value = "-2"
    
    Call BtnDisabled(1)
    Call LayerShowHide(1)
    
    objForm.Action = strUrl
    objForm.submit
    
    Call LayerShowHide(0)
    Call BtnDisabled(0)

End Sub



'========================================================================================
' Function Name : FncUsrEBRPreview(EBRName, condvar, x, y)
' Description   : Easybase web Preview  (EBR) - User Defined Size
'========================================================================================

Sub FncUsrEBRPreview(EBRName, condvar, x, y)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(5), arrField, arrHeader
	Dim pre

	If GetEBVer = "REQUBE" then
	    Call FncEBR5(EBRName, "view", CondVar, x, y, x, y)
	    Exit Sub	
	End If
	
	pre = GetPath(window.location.href)
	
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath
	strUrl = strUrl & "ExecuteWinReport"
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbDbName
	arrParam(2) = EBParent.gEbPkgRptPath & "\" & gLang & "\EBR\" & EBRName
	arrParam(3) = condvar & strPolicy
	arrParam(4) = "-2"
	arrParam(5) = strUrl
	
	Call BtnDisabled(1)	

	window.showModalDialog pre & "comasp/EBRpreview.asp?EBWidth=" & x & "&EBHeight=" & y &"""", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=" & x & "px; dialogHeight=" & y & "px; center: Yes; help: No; resizable: Yes; status: No; scroll:no;"

	Call BtnDisabled(0)	
	
End Sub

'========================================================================================
' Function Name : FncUsrEBCPreview(EBCName, condvar, x, y)
' Description   : Easybase web Preview  (EBC)  User Defined Size
'========================================================================================

Sub FncUsrEBCPreview(EBCName, condvar, x, y)

	Dim strUrl
	Dim strPolicy
	Dim arrParam(6), arrField, arrHeader
	Dim pre

	if GetEBVer = "REQUBE" then
        Call FncEBR5RC(EBCName, "view", condvar, x, y, x, y, "EBC")
	    Exit Sub	
	end if

	
	pre = GetPath(window.location.href)
	strPolicy = setPolicy()
	
	strUrl = strUrl & EBParent.gServerIP
	strUrl = strUrl & EBParent.gEbEnginePath
	strUrl = strUrl & "ExecuteWinCrossTab"
	
	arrParam(0) = EBParent.gEbUserName
	arrParam(1) = EBParent.gEbDbName
	arrParam(2) = EBParent.gEbPkgRptPath & "\" & gLang & "\EBC\" & EBCName
	arrParam(3) = condvar & strPolicy
	arrParam(4) = "-2"
	arrParam(5) = strUrl
	arrParam(6) = "1"

	Call BtnDisabled(1)		
	
	window.showModalDialog pre & "comasp/EBCpreview.asp?EBWidth=" & x & "&EBHeight=" & y &"""", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=" & x & "px; dialogHeight=" & y & "px; center: Yes; help: No; resizable: Yes; status: No; scroll:no;"

	Call BtnDisabled(0)	

End Sub

'========================================================================================
' Function Name : EBQuery(Byval pEbId,Byval pId,Byval pType,ObjName)
' Description   : Easybase DB Connection
'========================================================================================
Function AskEBDocumentName(Byval pEbId,Byval pType)
 	
	IntRetCD = CommonQueryRs("MNU_EB_CALL_NM,MNU_EB_TYPE","Z_DC_EBNAME","MNU_EB_ID = '"& pEbId & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 & "" = "" Then
		ObjName = pEbId+"."+pType
	Else
		lgF0 = Split(lgF0, chr(11))
		lgF1 = Split(lgF1, chr(11))
		If Trim(lgF0(0)) = "" Then
			ObjName = pEbId + "." + pType
		Else
			ObjName = Trim(lgF0(0)) + "." + Trim(lgF1(0))
		End If
	End If
	
	AskEBDocumentName = ObjName
End Function	


'========================================================================================
' Function Name : EBQuery(Byval pEbId,Byval pId,Byval pType,ObjName)
' Description   : Easybase DB Connection
'========================================================================================
Sub AskEBDocumentName2(Byval pEbId,prWidth,prHeight,prWidth2,prHeight2)
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim IntRetCD
 	
	IntRetCD = CommonQueryRs("isnull(form_width,1024),isnull(form_height,768),isnull(form_width2,1015),isnull(form_height2,715)","Z_DC_EBNAME","MNU_EB_ID = '"& pEbId & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    prWidth  = 1024
    prHeight = 768
    
	prWidth2 = 1015
	prHeight2 = 715

	If lgF0 & "" = "" Then
	Else

		lgF0 = Split(lgF0, chr(11))
		lgF1 = Split(lgF1, chr(11))
		
		prWidth  = lgF0(0)
		prHeight = lgF1(0)
		
		lgF2 = Split(lgF2, chr(11))
		lgF3 = Split(lgF3, chr(11))
		
		prWidth2  = lgF2(0)
		prHeight2 = lgF3(0)		
		

	End If

    If prWidth = 0 Then
       prWidth  = 1024
       prHeight = 768
	End If
	
    If prWidth2 = 0 Then
       prWidth2  = 1015
       prHeight2 = 715
	End If	
	
End Sub	
