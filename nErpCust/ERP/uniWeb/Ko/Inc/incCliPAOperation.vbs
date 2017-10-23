Dim M990012
Dim M990013
Dim M990014
Dim M990015

'======================================================================================================
Sub AdjustStyleSheet(pDoc)
    Call PopupParent.AdjustStyleSheet(pDoc)
End Sub

'======================================================================================================
Function CmpCharLength(ByVal szAllText,ByVal strLen) 
    CmpCharLength = PopupParent.CmpCharLength(szAllText,strLen)
End Function 

'======================================================================================================
Sub ElementVisible(objElement, ByVal Status)
    Call PopupParent.ElementVisible(objElement,Status)
End Sub

'======================================================================================================
Function GetSetupMod(ByVal strSetupMod, ByVal strCheckMod)
    GetSetupMod = PopupParent.GetSetupMod(strSetupMod,strCheckMod)
End function

'======================================================================================================
Sub ProtectTag(objName)
    Call PopupParent.ProtectTag(objName)
End Sub

'======================================================================================================
Sub ReleaseTag(objName)
    Call PopupParent.ReleaseTag(objName)	
End Sub

'======================================================================================================
Function IsBetween(ByVal iFrom,ByVal iTo,ByVal iIt)

    IsBetween =  False
    If iIt >= iFrom And iIt <= iTo Then
       IsBetween = True
    End If
End Function

'======================================================================================================
Function EnCoding(Byval iStr)
    EnCoding = iStr
End Function

'======================================================================================================
Function CountStrings(ByVal strString, ByVal strTarget)
     CountStrings = PopupParent.CountStrings(strString,strTarget)
End Function

Sub SetSpreadBackColor(pSpread, ByVal Row1,ByVal Col1,ByVal Row2,ByVal Col2,ByVal pvBackColor)
    pSpread.BlockMode = True
    pSpread.Row  = Row1
    pSpread.Col  = Col1
    pSpread.Row2 = Row2
    pSpread.Col2 = Col2
    pSpread.BackColor = pvBackColor
    pSpread.BlockMode = False
End Sub
'######################################################################################################
'
'
'
'
'  String Function List
'
'
'
'
'######################################################################################################

'======================================================================================================
Function ValueEscape(strURL)
    ValueEscape = PopupParent.ValueEscape(strURL)
End Function

'======================================================================================================
Function PreEscape(ByVal strVal)
	PreEscape = Escape(strVal)
End Function

'======================================================================================================
Function FilterVar(ByVal pStr,ByVal pStrALT,ByVal pOpt)
     FilterVar = PopupParent.FilterVar(pStr,pStrALT,pOpt)
End Function

'==============================================================================
Function ValidateData(ByVal pNum ,ByVal pOpt )
    ValidateData = PopupParent.ValidateData(pNum ,pOpt )
End Function  

'######################################################################################################
'
'
'
'
'  Spread Function List
'
'
'
'
'######################################################################################################

'========================================================================
Sub GoToCondition(pDoc)
    Call PopupParent.GoToCondition(pDoc)    
End Sub

Sub SetActiveCell(pvSpread,ByVal Col, ByVal Row,ByVal pScreenType,ByVal pDummy1,ByVal pDummy2)
    Call PopupParent.SetActiveCell(pvSpread,Col,Row,pScreenType,pDummy1,pDummy2)
    Set gActiveElement = document.activeElement
End Sub

'======================================================================================================
Function VisibleRowCnt(pDoc,ByVal pStartRow)
    VisibleRowCnt = PopupParent.VisibleRowCnt(pDoc,pStartRow)
End Function

'======================================================================================================
Function FncSumSheet(pObject,pPiVot,pStart,pEnd,pBool,pTargetRow,pTargetCol,pVerHor)
     FncSumSheet = PopupParent.FncSumSheet(pObject,pPiVot,pStart,pEnd,pBool,pTargetRow,pTargetCol,pVerHor)
End Function    

'======================================================================================================
function importExcel(objSpread)
	Dim arrRet
	Dim arrParam(1)
	Dim idx
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrParam(0) = objSpread.MaxCols    
	arrRet = window.showModalDialog("../../comasp/ImportExcel.asp", Array(arrParam), _
		"dialogWidth=450px; dialogHeight=130px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	objSpread.MaxRows = 0
	ggoSpread.Source = objSpread
	ggoSpread.SSShowData arrRet
	objSpread.col = 0
	For idx = 1 to objSpread.MaxRows
		objSpread.row   = idx
		objSpread.text = ggoSpread.InsertFlag                                      '☜: Insert
	Next
	
End Function

'======================================================================================================
Sub CheckMinNumSpread(pObject, ByVal Col, ByVal Row)
    Call PopupParent.CheckMinNumSpread(pObject, Col, Row)
End Sub

'######################################################################################################
'
'
'
'
'  Date Function List
'
'
'
'
'
'######################################################################################################

'======================================================================================================
Function UNIConvDate(ByVal pDate)
    UNIConvDate = PopupParent.UNIConvDate(pDate)
End Function

'======================================================================================================
Function UNIDateClientFormat(ByVal pDate)
	UNIDateClientFormat =  PopupParent.UNIDateClientFormatSub(pDate,"YMD")
End Function

'======================================================================================================
Function UNIMonthClientFormat(ByVal pDate)
    UNIMonthClientFormat = PopupParent.UNIDateClientFormatSub(pDate,"YM")
End Function

'======================================================================================================
Function UNIConvDateDBToCompany(ByVal pDate, ByVal pDefault)
    UNIConvDateDBToCompany = PopupParent.UNIConvDateDBToCompany(pDate, pDefault)
End Function

'======================================================================================================
Function UNIConvDateCompanyToDB(ByVal pDate, ByVal pDefault)
    UNIConvDateCompanyToDB = PopupParent.UNIConvDateCompanyToDB(pDate, pDefault)
End Function

'======================================================================================================
Function UNIFormatDate(Byval pDate)
    UNIFormatDate = PopupParent.UniConvLocalToCompanyDateFormat(pDate,"YMD")
End Function

'======================================================================================================
Function UNIFormatMonth(Byval pDate)
    UNIFormatMonth = PopupParent.UniConvLocalToCompanyDateFormat(pDate,"YM")
End Function

'======================================================================================================
Function UniConvLocalToCompanyDateFormat(ByVal pDate, ByVal pOpt)
    UniConvLocalToCompanyDateFormat = PopupParent.UniConvLocalToCompanyDateFormat(pDate, pOpt)
End Function

'======================================================================================================
Function UNICDate(ByVal pDate)
    UNICDate = PopupParent.UNICDate(pDate)	
End Function

'======================================================================================================
Function UniConvDateToYYYYMMDD(ByVal pDate , ByVal pDateFormat , ByVal pDateSeperator)
    UniConvDateToYYYYMMDD = PopupParent.UniConvDateToYYYYMMDD(pDate , pDateFormat , pDateSeperator)
End Function

'======================================================================================================
Function UniConvYYYYMMDDToDate(ByVal pDateFormat ,ByVal strYear,ByVal strMonth,ByVal strDay)
    UniConvYYYYMMDDToDate = PopupParent.UniConvYYYYMMDDToDate( pDateFormat , strYear, strMonth, strDay)
End Function

'======================================================================================================
Function UniConvDateToYYYYMM(ByVal pDate , ByVal pDateFormat , ByVal pDateSeperator)
    UniConvDateToYYYYMM = PopupParent.UniConvDateToYYYYMM( pDate ,  pDateFormat ,  pDateSeperator)
End Function

'======================================================================================================
Function UniConvDateAToB(ByVal pDate , ByVal pFromDateFormat , ByVal pToDateFormat)
    UniConvDateAToB = PopupParent.UniConvDateAToB(pDate , pFromDateFormat , pToDateFormat)
End Function

'======================================================================================================
Function UNIDateAdd(ByVal pInterVal , ByVal pNumber, ByVal pDate, ByVal pDateFormat)
    UNIDateAdd = PopupParent.UNIDateAdd(pInterVal ,pNumber,pDate,pDateFormat)
End Function

'======================================================================================================
Function UNIGetLastDay(ByVal pDate,ByVal pDateFormat)
    UNIGetLastDay = PopupParent.UNIGetLastDay(pDate,pDateFormat)
End Function

'======================================================================================================
Function UNIGetFirstDay(ByVal pDate,ByVal pDateFormat)
    UNIGetFirstDay =  PopupParent.UNIGetFirstDay(pDate,pDateFormat)
End Function

'==============================================================================
Function CheckDateFormat(ByVal pDate , ByVal pDateFormat)
     CheckDateFormat = PopupParent.CheckDateFormat( pDate ,  pDateFormat)
End Function

'==============================================================================
Sub ExtractDateFrom(ByVal pDate,pDateFormat,pDateSeperator,strYear,strMonth,strDay)
    Call PopupParent.ExtractDateFrom(pDate,pDateFormat,pDateSeperator,strYear,strMonth,strDay)
End Sub

'==============================================================================
Function MakeDateTo(pOpt,pDateFormat,pDateSeperator,strYear,strMonth,strDay)
    MakeDateTo =  PopupParent.MakeDateTo(pOpt,pDateFormat,pDateSeperator,strYear,strMonth,strDay)
End Function

'==============================================================================
Function FillLeadingSpaceWithZero(ByVal pDate,pDateFormat)
    FillLeadingSpaceWithZero = PopupParent.FillLeadingSpaceWithZero(pDate,pDateFormat)
End Function

'======================================================================================================
'
'
'
'
'  Numeric Function List
'
'
'
'
'======================================================================================================

'======================================================================================================
Function UNICCur(ByVal pNum)
    UNICCur = PopupParent.UNICCur(pNum)
End Function

'==============================================================================
Function UNIConvNum(ByVal pNum, ByVal pDefault)
    UNIConvNum = PopupParent.UNIConvNum(pNum,pDefault)
End Function

'======================================================================================================
Function UNICDbl(ByVal pNum)
    UNICDbl = PopupParent.UNICDbl(pNum)
End Function

'======================================================================================================
Function UniConvNumPCToCompanyWithoutRound(ByVal pNum,ByVal pDefault)
    UniConvNumPCToCompanyWithoutRound =  PopupParent.UniConvNumPCToCompanyWithoutRound(pNum,pDefault)
End Function

'======================================================================================================
Function UNIFormatNumber(ByVal pNum, ByVal pDecPoint, ByVal pFormatType, ByVal pNegativeNum,ByVal pRndPolicy, ByVal pRndUnit)
    UNIFormatNumber = PopupParent.UNIFormatNumber(pNum,pDecPoint, pFormatType, pNegativeNum,pRndPolicy,pRndUnit)
End Function

'========================================================================================
Function FncRoundData(ByVal pNum, ByVal pDecPoint, ByVal pRndPolicy, ByVal pRndUnit, ByVal pNumDec)
    FncRoundData = PopupParent.FncRoundData( pNum,  pDecPoint,  pRndPolicy,  pRndUnit,  pNumDec)
End Function

'=========================================================================================================================
Function uniConvNumAToB(ByVal pNum, ByVal pNum1000From, ByVal pNumDecFrom, ByVal pNum1000To, ByVal pNumDecTo, ByVal p1000SEP,ByVal pOpt1,ByVal pOpt2)
    uniConvNumAToB = PopupParent.uniConvNumAToB( pNum,  pNum1000From,  pNumDecFrom,  pNum1000To,  pNumDecTo,  p1000SEP, pOpt1, pOpt2)
End Function

'######################################################################################################
'
'
'
'
'  MA dependent Function List
'
'
'
'
'
'######################################################################################################

'========================================================================================
' Function Name : AppendNumberPlace
' Function Desc : Append Number place
'========================================================================================
Sub AppendNumberPlace(ByVal iiPos,ByVal iIntegeral,ByVal iDec)
    Dim iDx
    Dim sBuffer1
    Dim sBuffer2

    iiPos = CInt(iiPos)

    If iiPos >= 2 And iiPos <=5 Then
       Exit Sub
    End If
    
    If iiPos >= 10 Then
       Exit Sub
    End If
    
    If Trim(ggStrIntegeralPart) = "" Then 
       For iDx = 0 To 13 
          ggStrIntegeralPart = ggStrIntegeralPart & PopupParent.gColSep 
       Next   
    End If
    
    If Trim(ggStrDeciPointPart) = "" Then 
       For iDx = 0 To 13 
          ggStrDeciPointPart = ggStrDeciPointPart & PopupParent.gColSep 
       Next   
    End If

    ggStrIntegeralPart =  Split(ggStrIntegeralPart,PopupParent.gColSep)
    ggStrDeciPointPart =  Split(ggStrDeciPointPart,PopupParent.gColSep)
    
    If Trim(iIntegeral) = "" Then
       iIntegeral = 15 - CInt(iDec)
    End If
    
    ggStrIntegeralPart(iiPos) = CStr(iIntegeral)
    ggStrDeciPointPart(iiPos) = CStr(iDec)
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    For iDx = 0 To 13
        sBuffer1 =  sBuffer1 & ggStrIntegeralPart(iDx) & PopupParent.gColSep
        sBuffer2 =  sBuffer2 & ggStrDeciPointPart(iDx) & PopupParent.gColSep
    Next
    
    ggStrIntegeralPart = sBuffer1
    ggStrDeciPointPart = sBuffer2

End Sub

'===============================================================================
' Function Name : AppendNumberRange
' Function Desc : append number range
'===============================================================================
Sub AppendNumberRange(ByVal iPos,ByVal iMin,ByVal iMax)
    Dim iDx
    Dim iiPos
    Dim sBuffer1
    Dim sBuffer2

    iiPos = CInt(iPos)
    
    If Trim(ggStrMinPart) = "" Then 
       For iDx = 0 To 10 
          ggStrMinPart = ggStrMinPart & PopupParent.gColSep 
       Next   
    End If    '

    If Trim(ggStrMaxPart) = "" Then 
       For iDx = 0 To 10 
          ggStrMaxPart = ggStrMaxPart & PopupParent.gColSep 
       Next   
    End If       

    ggStrMinPart =  Split(ggStrMinPart,PopupParent.gColSep)
    ggStrMaxPart =  Split(ggStrMaxPart,PopupParent.gColSep)
    
    ggStrMinPart(iiPos) = CStr(iMin)
    ggStrMaxPart(iiPos) = CStr(iMax)
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    For iDx = 0 To 9
        sBuffer1 =  sBuffer1 & ggStrMinPart(iDx) & PopupParent.gColSep
        sBuffer2 =  sBuffer2 & ggStrMaxPart(iDx) & PopupParent.gColSep
    Next
    ggStrMinPart = sBuffer1
    ggStrMaxPart = sBuffer2

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
	
		If Ucase(objBtn.TagName) = "BUTTON" And objBtn.getAttribute("Flag") = 1 then
			objBtn.disabled = Status
		end if
	Next
	
	Set objBtn = Nothing
	
	If Err.Number = 0 Then Err.Clear				    			

End Sub 

'========================================================================
'
'========================================================================
Function ChkField(pDoc, ByVal pStrGrp)
    On Error Resume Next
    
    Dim i, intDivCnt, intTagNum
    Dim strTagName, strRequired
    Dim iRet
    Dim iRet2
    Dim iRequired
    
    iRequired = UCase(PopupParent.UCN_REQUIRED)
            
    intDivCnt = 0
    ChkField = False
    
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
                
        intTagNum = Mid(pDoc.All(i).Tag, 1, 1)
        strRequired = UCase(pDoc.All(i).className)
        
        
        If Err.Number <> 0 Then
            Err.Clear
        Else			
            If (intTagNum = pStrGrp Or pStrGrp = "A") And strRequired = iRequired Then
                Select Case strTagName
                    Case "INPUT", "TEXTAREA", "SELECT"
                        If Len(Trim(pDoc.All(i).Value)) = 0 Then
                            If intTagNum = "1" Then
                                iRet = DisplayMsgBox("970029", "X", pDoc.All(i).alt,"x")
                            Else
                                iRet = DisplayMsgBox("970021", "X", pDoc.All(i).alt,"x")
                            End If
                            
                            Call ChangeTabs2(pDoc, intDivCnt)
                            pDoc.All(i).focus
                            Set gActiveElement = document.activeElement       
                            Exit Function
                        End If
                        
                    Case "OBJECT"
                        If pDoc.All(i).Title = "FPDATETIME" Or pDoc.All(i).Title = "FPDOUBLESINGLE" Then
                            If Len(Trim(pDoc.All(i).Text)) = 0 Then
                                If intTagNum = "1" Then
                                   iRet = DisplayMsgBox("970029", "X", pDoc.All(i).alt,"x")
                                Else
                                   iRet = DisplayMsgBox("970021", "X", pDoc.All(i).alt,"x")
                                End If
                                
                                Call ChangeTabs2(pDoc, intDivCnt)                                                                
                             	Call SetFocusToDocument("P")
                                pDoc.All(i).focus
                                
                                Set gActiveElement = document.activeElement                            
                                Exit Function
                            End If
                        End If
                        
                End Select
                
            End If
            
        End If
        
    Next
    
    ChkField = ChkFieldLength(pDoc, pStrGrp)
    
End Function
  
'========================================================================
'
'========================================================================
Function ChkFieldLength(pDoc, ByVal pStrGrp)
    On Error Resume Next
    
    Dim i, intDivCnt, intTagNum
    Dim strTagName, strRequired
    Dim iRet
    Dim iMaxLen
    Dim iProtected 
    iProtected = UCase(PopupParent.UCN_PROTECTED)

    intDivCnt = 0
    ChkFieldLength = False
    
    
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
                
        intTagNum = Mid(pDoc.All(i).Tag, 1, 1)
        strRequired = UCase(pDoc.All(i).className)
        
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If (intTagNum = pStrGrp Or pStrGrp = "A") Then
                Select Case strTagName
                    Case "INPUT"
                        If UCase(pDoc.All(i).Type) = "TEXT" Then
                           iMaxLen = CDbl(pDoc.All(i).maxLength)
                           'If iMaxLen < 256 Then
                              If strRequired <> iProtected Then    
                                 If PopupParent.CmpCharLength(Trim(pDoc.All(i).Value),iMaxLen) = false Then
                                    iRet = DisplayMsgBox("900028", "X", pDoc.All(i).alt,"x")
                                    Call ChangeTabs2(pDoc, intDivCnt)
                                    pDoc.All(i).focus
                                    Set gActiveElement = document.activeElement                            
                                    Exit Function
                                 End If
                              End If
                           'End If
                        End If
                        
      '              Case "TEXTAREA"
       '                    If  CmpCharLength(Trim(pDoc.All(i).Value),128) = false Then
        '                       iRet = DisplayMsgBox("900028", "X", pDoc.All(i).alt,"x")
         '                      Call ChangeTabs2(pDoc, intDivCnt)
          '                     pDoc.All(i).focus
           '                    Set gActiveElement = document.activeElement                            
            '                   Exit Function
             '              End If
                End Select
                
            End If
            
        End If
        
    Next
    
    ChkFieldLength = True
    
End Function  

'=============================================================================
' Function Name  : ChangeTabs
'=============================================================================
Sub ChangeTabs2(ByRef objDoc, ByVal pPageNo)
    Dim gImgFolder

    Dim panel
    Dim myTabs
    
    Dim iPageNo
    
    Dim iLoc
    Dim strLoc

    If gPageNo = pPageNo Then 
       Exit Sub
    End If 
    
    Set panel = objDoc.All.TabDiv
    Set myTabs = objDoc.All.MyTab
    
    iPageNo = 0

    strLoc = objDoc.All.MyTab(pPageNo - 1).rows(0).cells(1).background
    iLoc = 1
    
    iLoc = InStrRev(strLoc, "/", -1)
    gImgFolder = Left(strLoc, iLoc)
    
    ' "../../image/table/tab_up_bg.gif"
    
    For iPageNo = 0 To panel.Length - 1
        
        myTabs(iPageNo).rows(0).cells(0).background = gImgFolder + "tab_up_bg.gif"
        myTabs(iPageNo).rows(0).cells(1).background = gImgFolder + "tab_up_bg.gif"
        myTabs(iPageNo).rows(0).cells(2).background = gImgFolder + "tab_up_bg.gif"'

        myTabs(iPageNo).rows(0).cells(0).children(0).src = gImgFolder + "tab_up_left.gif"
        myTabs(iPageNo).rows(0).cells(2).children(0).src = gImgFolder + "tab_up_right.gif"
       panel(iPageNo).Style.display = "none"
        
    Next
    
    ' 각각의 Tab 속성을 Default, Display None으로 설정 
    myTabs(pPageNo - 1).rows(0).cells(0).background = gImgFolder + "seltab_up_bg.gif"
    myTabs(pPageNo - 1).rows(0).cells(1).background = gImgFolder + "seltab_up_bg.gif"
    myTabs(pPageNo - 1).rows(0).cells(2).background = gImgFolder + "seltab_up_bg.gif"

    myTabs(pPageNo - 1).rows(0).cells(0).children(0).src = gImgFolder + "seltab_up_left.gif"
    myTabs(pPageNo - 1).rows(0).cells(2).children(0).src = gImgFolder + "seltab_up_right.gif"
    panel(pPageNo - 1).Style.display = ""
    
    gPageNo     = pPageNo
    
 End Sub
 
'======================================================================================================
' Function Name : CheckRunningBizProcess
' Function Desc : Check If Biz ASP is running
'======================================================================================================
Function CheckRunningBizProcess()

	CheckRunningBizProcess = True

	If window.document.all("MousePT").style.visibility = "visible" Then 
	   Exit Function
	End If   

	CheckRunningBizProcess = False

End Function

'===============================================================================
' Function Name : CompareDateByFormat
' Function Desc : Compare Date
'===============================================================================
Function CompareDateByFormat(pFromDt, pToDt,pFromDtAlt,pToDtAlt,ByVal pMsgCD,ByVal pDateFormat,ByVal pDateSeperator,pBool)

    Dim strYear1,strMonth1,strDay1,strFullDay1
    Dim strYear2,strMonth2,strDay2,strFullDay2

	CompareDateByFormat = False
    
    Call PopupParent.ExtractDateFrom(pFromDt,pDateFormat,pDateSeperator,strYear1,strMonth1,strDay1)
    Call PopupParent.ExtractDateFrom(pToDt  ,pDateFormat,pDateSeperator,strYear2,strMonth2,strDay2)
      
    strFullDay1 = strYear1 & strMonth1 & strDay1
    strFullDay2 = strYear2 & strMonth2 & strDay2

	If Len(Trim(strFullDay2)) Then
       If Len(Trim(strFullDay1)) Then
          If strFullDay1 > strFullDay2 Then
             If pBool = True Then
                If pMsgCD = "970023" Then
                   Call DisplayMsgBox(pMsgCD,"X", pToDtAlt, pFromDtAlt)
                Else
                   Call DisplayMsgBox(pMsgCD,"X", pFromDtAlt , pToDtAlt)
                End If   
             End If   
             Exit Function
          End If
       End If
	End If

	CompareDateByFormat = True

End Function

'===============================================================================
' Function Name : CompareDateByFormat
' Function Desc : Compare Date
'===============================================================================
Function CompareDateByFormat2(pFromDt, pToDt,pFromDtAlt,pToDtAlt,ByVal pMsgCD,ByVal pDateFormat,ByVal pDateSeperator,pBool)

    Dim strYear1,strMonth1,strDay1,strFullDay1
    Dim strYear2,strMonth2,strDay2,strFullDay2

	CompareDateByFormat2 = False
    
    Call PopupParent.ExtractDateFrom(pFromDt,pDateFormat,pDateSeperator,strYear1,strMonth1,strDay1)
    Call PopupParent.ExtractDateFrom(pToDt  ,pDateFormat,pDateSeperator,strYear2,strMonth2,strDay2)
      
    strFullDay1 = strYear1 & strMonth1 '& strDay1
    strFullDay2 = strYear2 & strMonth2 '& strDay2

	If Len(Trim(strFullDay2)) Then
       If Len(Trim(strFullDay1)) Then
          If (strFullDay1 + 11 ) < strFullDay2 Then
			Call DisplayMsgBox(pMsgCD,"X", pToDtAlt, pFromDtAlt)
            Exit Function
          End If
       End If
	End If

	CompareDateByFormat2 = True

End Function

'===============================================================================================
' Function Name : FindIndexOfCurrency
' Function Desc : This method initializes Type property as accounting to Currency
'===============================================================================================
Function FindIndexOfCurrency(ByVal pCurrency, ByVal pDataType)

	Dim iDx
	
    FindIndexOfCurrency = -1

	pCurrency = UCase(Trim(pCurrency))
	If pDataType = "3" Then
	   pCurrency = gCurrency
	End If   

	For iDx = 0 to UBound(gBCurrency)
		If UCASE(Trim(gBCurrency(iDx))) = pCurrency Then
           If gBDataType(iDx) = pDataType Then
              FindIndexOfCurrency = iDx
              Exit Function
           End If
        End If
	Next

End Function

'=================================================================================================================
'
'
'=================================================================================================================
Sub ReFormatSpreadCellByCellByCurrency(pObject,ByVal pStartRow,ByVal pEndRow,ByVal pCurrencyCol,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, ByVal Dummy1, ByVal Dummy2)
    Dim ii
    Dim iData
    Dim iCurrency
    Dim iDecimalPlaceAlignOpt
    Dim iDx
    Dim iArrDec, iDefaultDec, iDataType
    
    If UCase(pFormType) = "Q" Then
        iDecimalPlaceAlignOpt = PopupParent.gQMDPAlignOpt
    Else
        iDecimalPlaceAlignOpt = PopupParent.gIMDPAlignOpt
    End If

    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If

    If pObject.MaxRows = 0 Then
        Exit Sub
    End If

    If pStartRow = -1 Then
        pStartRow = 1
    End If

    If pEndRow = -1 Then
        pEndRow = pObject.MaxRows 
    End If

    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)
    iArrDec = Split(ggStrDeciPointPart,PopupParent.gColsep)
    iDefaultDec = iArrDec(iDataType + 8)
    For ii = pStartRow to pEndRow
        pObject.Col = pCurrencyCol
        pObject.Row = ii
        iCurrency = pObject.Text
        
        iDx = FindIndexOfCurrency(iCurrency,iDataType)
        pObject.Col = pTargetCol
        If iDx = -1 Then 
            pObject.TypeFloatDecimalPlaces = iDefaultDec      
        Else
            iData = PopupParent.UNICdbl(pObject.Text)
            pObject.Text = PopupParent.UNIFormatNumber(iData,gBDecimals(iDx) , -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx))
            If iDecimalPlaceAlignOpt = "1" Then
                pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
            Else
                pObject.TypeFloatDecimalPlaces = iDefaultDec
            End If
        End If
    Next
End Sub



'=================================================================================================================
'
'
'=================================================================================================================
Sub ReFormatSpreadCellByCellByCurrency2(pObject,ByVal pStartRow,ByVal pEndRow,ByVal pCurrency,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, ByVal Dummy1, ByVal Dummy2)
    Dim ii
    Dim iData
    Dim iDecimalPlaceAlignOpt
    Dim iDx
    Dim iArrDec, iDefaultDec, iDataType    
    
'    If UCase(pFormType) = "Q" Then
'        iDecimalPlaceAlignOpt = PopupParent.gQMDPAlignOpt
'    Else
'        iDecimalPlaceAlignOpt = PopupParent.gIMDPAlignOpt
'    End If

    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If

    If pObject.MaxRows = 0 Then
        Exit Sub
    End If

    If pStartRow = -1 Then
        pStartRow = 1
    End If

    If pEndRow = -1 Then
        pEndRow = pObject.MaxRows 
    End If
    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)    
    iArrDec = Split(ggStrDeciPointPart,PopupParent.gColsep)
    iDefaultDec = iArrDec(iDataType + 8)
    iDx = FindIndexOfCurrency(pCurrency,iDataType)
    For ii = pStartRow to pEndRow
        pObject.Row = ii
        pObject.Col = pTargetCol
        If iDx = -1 Then 
            pObject.TypeFloatDecimalPlaces = iDefaultDec        
        Else
            iData = PopupParent.UNICdbl(pObject.Text)
            pObject.Text = PopupParent.UNIFormatNumber(iData,gBDecimals(iDx) , -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx))
'            If iDecimalPlaceAlignOpt = "1" Then
                pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
'            Else
'                pObject.TypeFloatDecimalPlaces = iDefaultDec
'            End If
        End If
    Next
End Sub

'=================================================================================================================
'
'
'=================================================================================================================
Sub EditModeCheck(pObject,ByVal pRow,ByVal pCurrencyCol,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, Byval pMode, ByVal Dummy1, ByVal Dummy2)
    Dim iCurrency
    Dim iArrDec
    Dim iDecimalPlaceAlignOpt
    Dim iDx, iDataType
    
    If UCase(pFormType) = "Q" Then
        iDecimalPlaceAlignOpt = PopupParent.gQMDPAlignOpt
    Else
        iDecimalPlaceAlignOpt = PopupParent.gIMDPAlignOpt
    End If

    If iDecimalPlaceAlignOpt = "1" Then 
        Exit Sub
    End If
    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If

    If pObject.MaxRows = 0 Then
        Exit Sub
    End If
    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)    
    pObject.Row = pRow
    If pMode = 1 Then
        pObject.Col = pCurrencyCol
        iCurrency = pObject.Text
        
        iDx = FindIndexOfCurrency(iCurrency,iDataType)

        If iDx <> -1 Then
            pObject.Col = pTargetCol
            pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
        End If
    Else
        iArrDec = Split(ggStrDeciPointPart,PopupParent.gColsep)
        pObject.Col = pTargetCol
        pObject.TypeFloatDecimalPlaces = iArrDec(iDataType + 8)
    End If
End Sub
'=================================================================================================================
'
'
'=================================================================================================================
Sub EditModeCheck2(pObject,ByVal pRow,ByVal pCurrency,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, Byval pMode, ByVal Dummy1, ByVal Dummy2)
    Dim iArrDec
    Dim iDecimalPlaceAlignOpt
    Dim iDx, iDataType
    
'    If UCase(pFormType) = "Q" Then
'        iDecimalPlaceAlignOpt = PopupParent.gQMDPAlignOpt
'    Else
'        iDecimalPlaceAlignOpt = PopupParent.gIMDPAlignOpt
'    End If

'    If iDecimalPlaceAlignOpt = "1" Then 
        Exit Sub
'    End If
    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If

    If pObject.MaxRows = 0 Then
        Exit Sub
    End If
    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)    
    pObject.Row = pRow
    If pMode = 1 Then
        iDx = FindIndexOfCurrency(pCurrency,iDataType)
        
        If iDx <> -1 Then
            pObject.Col = pTargetCol
            pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
        End If
    Else
        iArrDec = Split(ggStrDeciPointPart,PopupParent.gColsep)
        pObject.Col = pTargetCol
        pObject.TypeFloatDecimalPlaces = iArrDec(iDataType + 8)
    End If
End Sub


'=================================================================================================================
'
'
'=================================================================================================================
Sub FixDecimalPlaceByCurrency(pObject,ByVal pRow, ByVal pCurrencyCol,ByVal pTargetCol,ByVal pDataType , ByVal Dummy1, ByVal Dummy2)
    Dim iData    
    Dim iCurrency
    Dim iTemp    
    Dim iArrDec
    Dim iDefaultDecimalPlace
    Dim iDx, iDataType

    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If
    
    If pObject.MaxRows = 0 Then
        Exit Sub
    End If
    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)    
    pObject.Col = pCurrencyCol
    pObject.Row = pRow
    iCurrency = pObject.Text
    
    iDx = FindIndexOfCurrency(iCurrency,iDataType)
    
    If iDx <> -1 Then
        pObject.Col = pTargetCol
        iTemp = 10 ^ gBDecimals(iDx)
        iData = Fix(CStr(PopupParent.UNICDbl(pObject.Text) * iTemp)) / iTemp
        pObject.Text = UniConvNumPCToCompanyWithoutRound(iData,"0")
    End If
End Sub

'=================================================================================================================
'
'
'=================================================================================================================
Sub FixDecimalPlaceByCurrency2(pObject,ByVal pRow, ByVal pCurrency,ByVal pTargetCol,ByVal pDataType , ByVal Dummy1, ByVal Dummy2)
    Dim iData    
    Dim iTemp    
    Dim iArrDec
    Dim iDefaultDecimalPlace
    Dim iDx, iDataType

    If UCase(TypeName(pObject)) = "EMPTY" Then
        Exit Sub
    End If
    
    If pObject.MaxRows = 0 Then
        Exit Sub
    End If
    iDataType = CStr(ASC(UCase(pDataType)) - ASC("A") + 2)    
    iDx = FindIndexOfCurrency(pCurrency,iDataType)
    
    pObject.Row = pRow
    If iDx <> -1 Then
        pObject.Col = pTargetCol
        iTemp = 10 ^ gBDecimals(iDx)
        iData = Fix(CStr(PopupParent.UNICDbl(pObject.Text) * iTemp)) / iTemp
        pObject.Text = UniConvNumPCToCompanyWithoutRound(iData,"0")
    End If
End Sub

'======================================================================================================
' Function Name : DisplayMsgBox
' Function Desc : 메세지 테이블을 검색하여 결과를 리턴한다.
'======================================================================================================
Function DisplayMsgBox(ByVal pMsgId,ByVal pBtnKind,ByVal pMsg1,ByVal pMsg2)

       DisplayMsgBox = ggoOper.DisplayMsgBox(pMsgId,pBtnKind,pMsg1,pMsg2)
    
End Function

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
	
End Sub

'========================================================================================
' Function Name : GetComaspFolderPath
' Description   : 
'========================================================================================
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

'==============================================================================
'
'==============================================================================
Sub SetFocusToDocument(pOpt)
   Select Case pOpt
     Case "M"
         top.Window.PopupParent.Frames(1).Focus
     Case "P"
         Window.Focus
   End Select      
End Sub

'======================================================================================================
' Function Name : SetCombo(pCombo, byval Code, byval Name)
' Parameters    : pCombo - Combo Object Name(SELECT Tag Name)		
'                 Code   - Code
'                 Name   - Text Value
'======================================================================================================
Sub SetCombo(pCombo, ByVal strValue, ByVal strText)
	Dim objEl
			
	Set objEl = Document.CreateElement("OPTION")	
	objEl.Text = strText
	objEl.Value = strValue

	pcombo.Add(objEl)
	Set objEl = Nothing

End Sub

'======================================================================================================
' Function Name : SetCombo2(pCombo, ByVal pCodeArr, ByVal pNameArr,pSeperator)
' Parameters    : pCombo - Combo Object Name(SELECT Tag Name)		
'                 Code   - Code
'                 Name   - Text Value
'======================================================================================================
Sub SetCombo2(pCombo, ByVal pCodeArr, ByVal pNameArr,pSeperator)

    Dim iDx

    pCodeArr = Split(pCodeArr,pSeperator)
    pNameArr = Split(pNameArr,pSeperator)
    
    For iDx = 0 To UBound(pCodeArr) - 1
        Call SetCombo(pCombo,pCodeArr(iDx), pNameArr(iDx))
    Next

End Sub

'======================================================================================================
' Function Name : SetCombo2(pCombo, ByVal pCodeArr, ByVal pNameArr,pSeperator)
' Parameters    : pCombo - Combo Object Name(SELECT Tag Name)		
'                 Code   - Code
'                 Name   - Text Value
'======================================================================================================
Sub SetCombo3(pCombo,  pCodeArr)

    Dim iLoop
    Dim iMax
    Dim iTemp,iTemp1
    
    if Trim(pCodeArr) = "" Then
       Exit Sub
    End If

    iTemp = Split(pCodeArr,Chr(12))
    
    iMax = UBound(iTemp)
    
    For iLoop = 0 To iMax - 1

        iTemp1 = Split(iTemp(iLoop),Chr(11))
        Call SetCombo(pCombo,iTemp1(0), iTemp1(1))
    Next

End Sub

'======================================================================================================
' Function Name : SpreadSetFloat
' Function Desc : This function is sets the type of float cells
'======================================================================================================
Sub SetSpreadFloat(ByVal iCol ,ByVal Header ,ByVal dColWidth ,ByVal HAlign ,ByVal iFlag )
    ggoSpread.SSSetFloat iCol,Header,dColWidth,CStr(iFlag),ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,HAlign
End Sub

'===============================================================================================
' Function Name : UNIFormatNumberByCurrecny
' Function Desc : This method initializes Numeric Format property as accounting to Currency
'               : Local Client System Number Type -> Company Informat Number Type
'===============================================================================================
Function UNIFormatNumberByCurrecny(ByVal pNum,ByVal pCurrency,ByVal pDataType)
    UNIFormatNumberByCurrecny = UNIConvNumPCToCompanyByCurrency(pNum, pCurrency, pDataType,"X", "X")    
End Function

'===============================================================================================
'   Function Name : uniFormatNumberByTax
'   Function Desc : This function initializes Numeric Format, which converts to Tax Format according to 
'                    tax rounding policy and amount decimal policy 
'                 : Local Client System Number Type -> Company Informat Number Type 
'===============================================================================================
Function uniFormatNumberByTax(ByVal pNum,ByVal pCurrency,ByVal pDataType)

    If PopupParent.ValidateData(pDataType,"SEN") = False Then
       pDataType = PopupParent.ggAmtOfMoneyNo
	End If
	
    uniFormatNumberByTax = UNIConvNumPCToCompanyByCurrency(pNum, pCurrency, pDataType,PopupParent.gTaxRndPolicyNo, "X")

End function

'===============================================================================================
' Function Name : UNIConvNumPCToCompanyByCurrency
' Function Desc :
'===============================================================================================
Function UNIConvNumPCToCompanyByCurrency(ByVal pNum,ByVal pCurrency,ByVal pDataType,ByVal pOpt1,ByVal pOpt2)

    Dim iDx
    Dim iRet
	
    UNIConvNumPCToCompanyByCurrency = ""
   
    iDx = FindIndexOfCurrency(pCurrency,pDataType)
	
    If CInt(iDx) < 0 Then 
       iDx = FindIndexOfCurrency(PopupParent.gCurrency,pDataType)

       If CInt(iDx) < 0 Then 
          iRet = MsgBox ("화폐별 포맷정보를 찾을 수가 없습니다." ,vbExclamation,PopupParent.gLogoName)  '2002/08/13 lee jinsoo
          UNIConvNumPCToCompanyByCurrency = PopupParent.UniConvNumPCToCompanyWithoutRound(pNum,"")       '2002/08/13 lee jinsoo
          Exit Function
        End If   
    End If
	
    Select Case pOpt1
       Case PopupParent.gTaxRndPolicyNo :   UNIConvNumPCToCompanyByCurrency = PopupParent.UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,PopupParent.gTaxRndPolicy  ,gBRoundingUnit(iDx))
       Case PopupParent.gLocRndPolicyNo : 
                                    If PopupParent.gBConfMinorCD = "1" Then
                                       UNIConvNumPCToCompanyByCurrency = PopupParent.UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gBRoundingPolicy(iDx) ,gBRoundingUnit(iDx)) 
                                    Else
                                       UNIConvNumPCToCompanyByCurrency = PopupParent.UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,PopupParent.gLocRndPolicy  ,gBRoundingUnit(iDx))
                                    End If   
       Case Else                  :    UNIConvNumPCToCompanyByCurrency = PopupParent.UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gBRoundingPolicy(iDx) ,gBRoundingUnit(iDx))
   End Select      

End Function


'===============================================================================
' Function Name : ValidDateCheck
' Function Desc : Valid Date Check Function
'===============================================================================
Function ValidDateCheck(pObjFromDt, pObjToDt)

	ValidDateCheck = False
	If Len(Trim(pObjToDt.Text)) Then
       If Len(Trim(pObjFromDt.Text)) Then
          If PopupParent.UniConvDateToYYYYMMDD(pObjFromDt.Text,pObjFromDt.UserDefinedFormat,"") > PopupParent.UniConvDateToYYYYMMDD(pObjToDt.Text,pObjToDt.UserDefinedFormat,"") Then
             Call DisplayMsgBox("970023","X", pObjToDt.Alt, pObjFromDt.Alt)
             Call SetFocusToDocument("P")
             pObjToDt.focus
             Set gActiveElement = document.activeElement
             Exit Function
          End If
       End If
	End If

	ValidDateCheck = True

End Function

'========================================================================================================
' 
' 
' 
'   Cli MA dependency
' 
' 
' 
' 
'========================================================================================================

'========================================================================================
' Function Name : AskSpdSheetAddRowCount
' Description   : 
'========================================================================================
Function AskSpdSheetAddRowCount()					' 2002-11-11 컬럼이동관련 추가 (김인태)
    Dim iRowCount
    Dim iRet
        
    AskSpdSheetAddRowCount = ""
        
    iRowCount = inputbox("추가할 행 수를 입력하세요.","행추가","1")
    If isEmpty(iRowCount) Then
        Exit Function
    End If
    iRowCount = Trim(iRowCount)
    If IsNumeric(iRowCount) And Len(iRowCount) < 5 Then
        iRowCount = CInt(iRowCount)
        If iRowCount >= 1 And iRowCount <= 999 Then         
            AskSpdSheetAddRowCount = iRowCount
            Exit Function
        End If
    End If
    
    iRet = MsgBox("1 이상 999 이하의 정수를 입력하세요.", vbInformation + vbQuestion, PopupParent.gLogoName)
        
End Function    

'========================================================================================
' Function Name : AskSpdSheetColumnName
' Description   : 
'========================================================================================
Function AskSpdSheetColumnName(Byval iColumnName)			' 2002-11-11 컬럼이동관련 추가 (김인태)
    iColumnName=Inputbox("컬럼의 타이틀을 입력하세요.","타이틀명변경",iColumnName)

    AskSpdSheetColumnName = Trim(iColumnName)
        
End Function

'========================================================================================================
'
'========================================================================================================
Sub PopChangeSpreadColumnname()
    Dim iColumnName

    If UCase(TypeName(gActiveSpdSheet)) = "EMPTY" Then
       Exit Sub
    End If
    
    If gActiveSpdSheet.MaxRows = 0 Then
       Exit Sub
    End If

    gActiveSpdSheet.Row = 0
    gActiveSpdSheet.Col = gActiveSpdSheet.ActiveCol
    iColumnName = gActiveSpdSheet.Text

    iColumnName = AskSpdSheetColumnName(iColumnName)

    If iColumnName <> "" Then
       ggoSpread.Source = gActiveSpdSheet
       Call ggoSpread.SSSetReNameHeader(gActiveSpdSheet.ActiveCol,iColumnName)
    End If
    
End Sub

'========================================================================================
' Function Name : PopSortPopup
' Description   : 
'========================================================================================
Function PopSortPopup()								' 2002-11-11 컬럼이동관련 추가 (김인태)
	Dim arrRet
	Dim arrParam
	Dim TInf(5)	
	Dim ii
	Dim iSortCol
	Dim iSortOrder
	Dim iSortFieldCD
	Dim iSortFieldNm
	Dim iPopUpR
    Dim iTempCount
    Dim iHiddenType
    Dim iCellTypeTemp
    
	On Error Resume Next

    If TypeName(gActiveSpdSheet) = "Empty" Then
       Exit Function
    End If

    ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )
    ReDim iSortCol(PopupParent.C_MaxSelList  - 1 )
    ReDim iSortOrder(PopupParent.C_MaxSelList  - 1 )
    ReDim iPopUpR(PopupParent.C_MaxSelList - 1,1)	

    TInf(0) = "정렬"

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSGetSortField(iPopUpR)
	Call ggoSpread.GetHiddenCol(iHiddenType)
    
    ReDim iSortFieldCD(ggoSpread.Source.MaxCols-1)
    ReDim iSortFieldNm(ggoSpread.Source.MaxCols-1)
    iTempCount = 0
    For ii = 1 To ggoSpread.Source.MaxCols-1
       ggoSpread.Source.Col = ii
       ggoSpread.Source.Row = -1
       iCellTypeTemp = ggoSpread.Source.CellType
       ggoSpread.Source.Col = ii
       ggoSpread.Source.Row = 0
       If iHiddenType(ii) <> 1 and iCellTypeTemp <> PopupParent.CT_BUTTON Then
          If iCellTypeTemp <> PopupParent.CT_CHECKBOX or Trim(ggoSpread.Source.Text) <> "" Then
             iSortFieldNm(iTempCount) = ggoSpread.Source.Text
             iSortFieldCD(iTempCount) = CStr(ii)
             iTempCount = iTempCount + 1
          End If
       End If
    Next
    ReDim preserve iSortFieldCD(iTempCount)  
    ReDim preserve iSortFieldNm(iTempCount)
        
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = iPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = iPopUpR(ii / 2  , 1)
    Next  

	arrRet = window.showModalDialog(GetComaspFolderPath & "CommonSortPopup.asp",Array(iSortFieldCD,iSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "0" Then
       If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
        ReDim iPopUpR(arrRet(0) / 2 - 1,1)
        ReDim iSortCol(arrRet(0) / 2 - 1)
        ReDim iSortOrder(arrRet(0) / 2 - 1)
        For ii = 0 to arrRet(0) - 1 Step 2
            iSortCol(ii / 2) = arrRet(ii + 1)  
            iSortOrder(ii / 2) = arrRet(ii + 2)
            iPopUpR(ii / 2 ,0) = CInt(arrRet(ii + 1))
            iPopUpR(ii / 2 ,1) = CInt(arrRet(ii + 2))
        Next 
        Call ggoSpread.SSSort2(iSortCol,iSortOrder)
        Call ggoSpread.SSSetSortField(iPopUpR)
    End If
End Function

'========================================================================================================
' Name : PopMakeHiddenColumn
' Desc : 
'========================================================================================================
Sub PopMakeHiddenColumn(ByVal Index,ByVal pTrueFalse)

    If IsNull(Index) Or Not IsNumeric(Index) Then
       Exit Sub 
    End If
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetColHidden(Index,Index,pTrueFalse,"D")    	    
End Sub

'========================================================================================================
' Name : PopUnfixCol
' Desc : 
'========================================================================================================
Sub PopUnfixCol()
    ggoSpread.Source = gActiveSpdSheet    
    Call ggoSpread.SSSetSplit(0)
End Sub

'========================================================================================
' Function Name : SetPopupMenuItemInf
' Description   : 
'========================================================================================
Sub SetPopupMenuItemInf(ByVal pPopupMenuItemBitInf)
    gPopupMenuItemBitInf = pPopupMenuItemBitInf
End Sub

'========================================================================================
' Function Name : EBQuery(Byval pEbId,Byval pId,Byval pType,ObjName)
' Description   : Easybase DB Connection
'========================================================================================
Function AskPRAspName(Byval pPgmId) 
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim IntRetCD
    
    AskPRAspName = ""
    
	IntRetCD = CommonQueryRs("pgm_nm,called_upper_fid,called_id","Z_PR_ASPNAME"," Lang_cd = '"& PopupParent.gLang & "' and pgm_id='"& pPgmId  &"'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD=False then
	Else
		lgF0 = split(lgF0,PopupParent.gColSep)
		lgF1 = split(lgF1,PopupParent.gColSep)
		lgF2 = split(lgF2,PopupParent.gColSep)
        PopupParent.gActivePRAspName = Trim(lgF0(0))
        
        lgF1(0) = UCase(Trim(lgF1(0)))
        lgF2(0) = UCase(Trim(lgF2(0)))
		If lgF1(0) = "COMASP" Then
		   AskPRAspName = GetRootFolderByLang & Trim(lgF1(0)) & "/" & Trim(lgF2(0)) & ".asp"
		Else
		   AskPRAspName = GetRootFolderByLang & "Module/" & Trim(lgF1(0)) & "/" & Trim(lgF2(0)) & ".asp"
		End If   
	End If
	
End Function

Function GetSpreadText(pSPD,ByVal Col,ByVal Row,ByVal pvDummy1,ByVal pvDummy2)   '2003/05/23
         Dim iRetBool
         GetSpreadText = ""
'         iRetBool = pSPD.GetText(Col,Row,GetSpreadText)		 ' 2003/06/09  float column
         pSPD.Col       = Col
         pSPD.Row       = Row
         GetSpreadText = pSPD.Text
End Function         

Function GetSpreadValue(pSPD,ByVal Col,ByVal Row,ByVal pvDummy1,ByVal pvDummy2)   '2003/05/23
         Dim iRetBool
         pSPD.Col       = Col
         pSPD.Row       = Row
         GetSpreadValue = pSPD.value
End Function    

Sub SetSpreadValue(pSPD,ByVal Col,ByVal Row,ByVal pValue,ByVal pvDummy1,ByVal pvDummy2)   '2003/05/23
    Dim iRetBool
    pSPD.Col   = Col
    pSPD.Row   = Row
    pSPD.value = pValue
End Sub

Sub CopySpreadValueAToB(pSPD,ByVal pCol1,ByVal pRow1,ByVal pCol2,ByVal pRow2,ByVal pvDummy1,ByVal pvDummy2)   '2003/05/23
    Dim intRow 
    Dim iTemp

    For intRow = pRow1 To pRow2
        pSPD.Row   = intRow 
        pSPD.Col   = pCol1
        iTemp      = pSPD.Value             ' .Value means that it is index of cell,not value in combo cell type
        pSPD.Row   = intRow 
        pSPD.Col   = pCol2
        pSPD.Value = iTemp
    Next
End Sub

'=======================================================================================
Public Sub FormatDATEField(TmpObject)

       TmpObject.ReDraw = False
       TmpObject.AlignTextV = 0
       TmpObject.AlignTextH = 1
       TmpObject.AllowNull = True
       TmpObject.Appearance = 1
       TmpObject.DateCalcMethod = 4  'Y2K
       TmpObject.DateCalcY2KSplit = 50
       TmpObject.InvalidOption = 2   'Clear Data
       TmpObject.DateTimeFormat = 5
       TmpObject.FontName = PopupParent.gFontName
       TmpObject.FontSize = PopupParent.gFontSize
       TmpObject.UserDefinedFormat = parent.gDateFormat
       TmpObject.DateDefault = "20000101"
       TmpObject.DateMin = "19000102"
       TmpObject.DateMax = "29991231"
       TmpObject.UserEntry = 0
       TmpObject.Value = ""
       TmpObject.ReDraw = True
       TmpObject.ButtonStyle = 1
       TmpObject.BorderColor = &H708090
                        
End Sub

'========================================================================================
Function ExternalWrite(strData)
	Document.Write strData
End Function

'===============================================================================================
' Ex)
' 
' Function GetAuth()
'
'    Dim xmlDoc
'    Dim idata_biz_area_cd
'    If GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) = True Then
'       idata_biz_area_cd = xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
'    End If
'       
' End Function
'
'data structure
'
'  <root>
'     <data_yn>
'       <data_biz_area_cd_yn>    </data_biz_area_cd_yn>
'       <data_internal_cd_yn>    </data_internal_cd_yn>
'       <data_sub_internal_cd_yn></data_sub_internal_cd_yn>
'       <data_personal_yn>       </data_personal_yn>
'       <data_plant_cd_yn>       </data_plant_cd_yn>
'       <data_pur_org_yn>        </data_pur_org_yn>
'       <data_pur_grp_yn>        </data_pur_grp_yn>
'       <data_sales_org_yn>      </data_sales_org_yn>
'       <data_sales_grp_yn>      </data_sales_grp_yn>
'       <data_sl_cd_yn>          </data_sl_cd_yn>
'       <data_wc_cd_yn>          </data_wc_cd_yn>
'     </data_yn>
'     <data>
'       <data_biz_area_cd_all>    </data_biz_area_cd_all>
'       <data_biz_area_cd>        </data_biz_area_cd>
'       <data_biz_area_nm>        </data_biz_area_nm>
'       <data_internal_cd_all>    </data_internal_cd_all>
'       <data_internal_cd>        </data_internal_cd>
'       <data_internal_nm>        </data_internal_nm>
'       <data_sub_internal_cd_all></data_sub_internal_cd_all>
'       <data_sub_internal_cd>    </data_sub_internal_cd>
'       <data_sub_internal_nm>    </data_sub_internal_nm>
'       <data_personal_id_all>    </data_personal_id_all>
'       <data_personal_id>        </data_personal_id>
'       <data_personal_nm>        </data_personal_nm>
'       <data_plant_cd_all>       </data_plant_cd_all>
'       <data_plant_cd>           </data_plant_cd>
'       <data_plant_nm>           </data_plant_nm>
'       <data_pur_grp_all>        </data_pur_grp_all>
'       <data_pur_grp>            </data_pur_grp>
'       <data_pur_grp_nm>         </data_pur_grp_nm>
'       <data_pur_org_all>        </data_pur_org_all>
'       <data_pur_org>            </data_pur_org>
'       <data_pur_org_nm>         </data_pur_org_nm>
'       <data_sales_org_all>      </data_sales_org_all>
'       <data_sales_org>          </data_sales_org>
'       <data_sales_org_nm>       </data_sales_org_nm>
'       <data_sales_grp_all>      </data_sales_grp_all>
'       <data_sales_grp>          </data_sales_grp>
'       <data_sales_grp_nm>       </data_sales_grp_nm>
'       <data_sl_cd_all>          </data_sl_cd_all>
'       <data_sl_cd>              </data_sl_cd>
'       <data_sl_nm>              </data_sl_nm>
'       <data_wc_cd_all>          </data_wc_cd_all>
'       <data_wc_cd>              </data_wc_cd>
'       <data_wc_nm>              </data_wc_nm>
'     </data>
'  </root>
'
'===============================================================================================
Function GetDataAuthXML(ByVal pUID, ByVal pMNUID, xmlDoc)
    Dim iTemp
    Dim iXmlHttp
    
    On Error Resume Next
    
    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP.3.0")
    
    GetDataAuthXML = False

    iXmlHttp.open "POST", GetComaspFolderPath & "SQLXMLProd.asp", False
    
    iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    iXmlHttp.send "LangCD=" & gLang & "&ADODBConnString=" & Escape(gADODBConnString) & "&StrSQL=" & Escape(" select dbo.ufn_z_get_mnu_auth_data('" & pUID & "','" & pMNUID & "') ")
    
    Set xmlDoc = iXmlHttp.responseXML
    
    
	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_biz_area_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_biz_area_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 사업장"
'            call history.go(-1)

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   


	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_internal_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_internal_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 내부부서"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   
	   
	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_sub_internal_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 내부부서(하위포함)"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   


	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_personal_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_personal_id_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_personal_id").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 개인"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   


	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_plant_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_plant_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_plant_cd").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 공장"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   


	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_pur_grp_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_pur_grp_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_pur_grp").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 구매그룹"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   



	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_pur_org_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_pur_org_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_pur_org").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 구매조직"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   



	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_sales_org_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_sales_org_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_sales_org").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 영업조직"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   


	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_sales_grp_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_sales_grp_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_sales_grp").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 영업그룹"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   



	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_sl_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_sl_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_sl_cd_yn").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 창고"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   



	if UCase(xmlDoc.selectSingleNode("/root/data_yn/data_wc_cd_yn").Text)     = "Y" Then
	   if UCase(xmlDoc.selectSingleNode("/root/data/data_wc_cd_all").Text) = "Y" then
	   else
	      if xmlDoc.selectSingleNode("/root/data/data_wc_cd").Text     = "" then
             MsgBox "권한설정 자료가 없습니다." & vbCrLf & vbCrLf & "항목 : 작업장"

             if InStr(LCase(document.location.href),"/module/") > 0 then
                document.location.href= "../../autherror.asp"
             else   
                document.location.href= "../autherror.asp"
             end if

             exit function
	      end if
	   end if
    end if   

    

    Set iXmlHttp = Nothing
    
    If iXmlHttp.parseError.errorCode = 0 Then
       GetDataAuthXML = True
    Else
       MsgBox xmlDoc.parseError.reason
    End If


End Function
