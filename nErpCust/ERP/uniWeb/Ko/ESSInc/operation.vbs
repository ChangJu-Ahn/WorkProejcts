Dim M990012
Dim M990013
Dim M990014
Dim M990015

'======================================================================================================
Sub AdjustStyleSheet(pDoc)

    On Error Resume Next
    
    Dim i
    
    For i = 0 To pDoc.All.Length - 1
        
        If UCase(pDoc.All(i).tagName) = "INPUT" Then
           If UCase(pDoc.All(i).TYPE) = "TEXT" Then
              If Mid(pDoc.All(i).getAttribute("tag"),6,1) = "U" Then
                 pDoc.All(i).style.textTransform = "uppercase"  
              End If   
           End If   
        End If   
          
    Next
    
End Sub

'=============================================================================
Function CmpCharLength(ByVal szAllText,ByVal strLen) 
    Dim nLen 
    Dim nCnt 
    Dim szEach 

    nLen = 0 
    szAllText = Trim(szAllText) 
    For nCnt = 1 To Len(szAllText) 
        szEach = Mid(szAllText,nCnt,1) 
        If 0 <= Asc(szEach) And Asc(szEach) <= 255 Then 
           nLen = nLen + 1             '한글이 아닌 경우 
        Else 
           nLen = nLen + 2             '한글인 경우 
        End If 
   Next 
   
   If nLen <= strLen Then
      CmpCharLength =  True
   Else
      CmpCharLength =  False
   End If

End Function 

'========================================================================================
Sub ElementVisible(objElement, ByVal Status)
	If Status = 0 Then 
		Status = "hidden"
	Else
		Status = "visible"
	End If
	objElement.style.visibility = Status
End Sub

'======================================================================================================
Function GetSetupMod(ByVal strSetupMod, ByVal strCheckMod)
  If instr(1, UCase(strSetupMod), UCase(strCheckMod)) > 0 Then
     GetSetupMod = "Y"
  Else
     GetSetupMod = "N"
  End if
End function

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
              objName.className = "form02"
              objName.disabled = "true"
          Case "RADIO"
              objName.className = "form02"
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

End Sub

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
Function MessageSplit( ByVal iCount, ByVal  MsgText, ByVal pMsg1, ByVal pMsg2)
  
     MsgText =  Replace(MsgText,"%1",pMsg1)
     MsgText =  Replace(MsgText,"%2",pMsg2)

     MessageSplit = MsgText
    
End Function

'======================================================================================================
Function CountStrings(ByVal strString, ByVal strTarget)
    Dim lPosition
    Dim iCount
   
    lPosition = 1
    
    If Trim(strString) = "" Then
       CountStrings = 0
       Exit Function
    End If
    
    Do While InStr(lPosition, strString, strTarget)
    
        lPosition = InStr(lPosition, strString, strTarget) + 1
        iCount = iCount + 1
    
    Loop    
    
    CountStrings = iCount
   
End Function

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
Function ConvSPChars(ByVal strVal)
	ConvSPChars = Replace(strVal, """", """""")
End Function 

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

'========================================================================================
' Trim string and set string to space if string length is zero
' pData   : target data
' pStrALT : alternative string if space
' pOpt    :  S is for String
'            D is for Digit
' History : Appended in 2002/08/07 (lee jin soo)
'========================================================================================
Function FilterVar(ByVal pData, ByVal pStrALT, ByVal pOpt)

     If IsNull(pData) Then
        pData = "" 
     Else   
        pData = Trim(pData)
     End If       
     
     pOpt = UCase(pOpt)
     
     Select Case VarType(pData)
        Case vbEmpty                                           '0    Empty (uninitialized)
                 FilterVar = pStrALT
                 Exit Function
        Case vbNull                                            '1    Null (no valid data)
                 FilterVar = "Null"
                 Exit Function
        Case vbInteger, vbLong, vbSingle, vbDouble             '2(Integer),3(Long integer),4(Single-precision floating-point number),5(Double-precision floating-point number)
                 FilterVar = pData
                 Exit Function
        Case vbCurrency, vbBoolean, vbByte                     '6(Currency),11(Boolean),17(Byte)
                 FilterVar = pData
                 Exit Function
        Case Else
        
                 If pData = "" Then
                    
                    If pOpt = "S" And Trim(pStrALT) = "" Then
                       pStrALT = "''"
                    End If
                    
                    If pOpt = "S2" And Trim(pStrALT) = "" Then
                       pStrALT = "''''"
                    End If
                    
                    If gCharSQLSet = "U" Then
                       If Len(pStrALT) > 1 Then
                          If Mid(pStrALT, 1, 2) = "N'" Then
                             pStrALT = Mid(pStrALT, 2)
                          End If
                       End If
                       
                       If pOpt = "S" Then
                       
                          If IsNull(pStrALT) Or UCase(Trim(pStrALT)) = "NULL" Then
                          Else
                             pStrALT = "N" & pStrALT
                          End If
                       
                       End If
                    
                    End If
                    
                    FilterVar = pStrALT
                    
                    Exit Function
                 End If
     
                 Select Case pOpt
                     Case "S"
                                pData = Replace(pData, "'", "''")
                                If gCharSQLSet = "U" Then
                                   FilterVar = "N'" & pData & "'"
                                Else
                                   FilterVar = "'" & pData & "'"
                                End If
                     Case "S2"
                                pData = Replace(pData, "'", "''")
                                If gCharSQLSet = "U" Then
                                   FilterVar = "N''" & pData & "''"
                                Else
                                   FilterVar = "''" & pData & "''"
                                End If
                     Case "SNM"
                                FilterVar = Replace(pData, "'", "''")
                     Case Else
                                FilterVar = pData
                 End Select
     End Select
     
End Function

'==============================================================================
Function ValidateData(ByVal pNum ,ByVal pOpt )

    ValidateData = False 
    
    If InStr(pOpt,"E") > 0 Then     
       If IsEmpty(pNum) Then
          Exit Function
       End If
    End If   

    If InStr(pOpt,"N") > 0 Then     
       If IsNull(pNum) Then
          Exit Function
       End If
    End If   

    If InStr(pOpt,"S") > 0 Then     
       If Trim(pNum) = "" Then
          Exit Function
       End If
    End If   
    
    If InStr(pOpt,"F") > 0 Then
       If Not IsNumeric(pNum) Then
          Exit Function
       End If    
    End If    
    
    ValidateData = True
      
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

Sub SetSpreadBackColor(pSpread, ByVal Row1,ByVal Col1,ByVal Row2,ByVal Col2,ByVal pvBackColor)
    pSpread.BlockMode = True
    pSpread.Row  = Row1
    pSpread.Col  = Col1
    pSpread.Row2 = Row2
    pSpread.Col2 = Col2
    pSpread.BackColor = pvBackColor
    pSpread.BlockMode = False
End Sub

'========================================================================
Sub GoToCondition(pDoc)
    On Error Resume Next
    
    Dim strTagName, strRequired
    Dim iRequired
    Dim i
    
    iRequired = UCase(UCN_PROTECTED)
            
    For i = 0 To pDoc.All.Length - 1
        
        strTagName = UCase(pDoc.All(i).tagName)
        
        strRequired = UCase(pDoc.All(i).className)
        
        
         If strRequired <> iRequired Then
                Select Case strTagName
                    Case "INPUT", "TEXTAREA", "SELECT"
                          pDoc.All(i).focus
                          Exit Sub
                    Case "OBJECT"
                        If pDoc.All(i).Title = "FPDATETIME" Or pDoc.All(i).Title = "FPDOUBLESINGLE" Then
                           pDoc.All(i).focus
                           Exit Sub
                        End If
                        
                End Select
                
         End If
        
    Next
    
End Sub

'======================================================================================================
Sub SetActiveCell(pvSpread,ByVal Col, ByVal Row,ByVal pScreenType,ByVal pDummy1,ByVal pDummy2)

    Call SetFocusToDocument(pScreenType)
       
    pvSpread.Focus
    pvSpread.Row    = Row
    pvSpread.Col    = Col
    pvSpread.Action = 0
    Set gActiveElement = document.activeElement
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
Function FncSumSheet(pObject,pPiVot,pStart,pEnd,pBool,pTargetRow,pTargetCol,pVerHor)
    Dim iDx
    Dim iSum, iSumTemp
    Dim iOperStatus
    
    iOperStatus =  True
    
    If pVerHor = "V" Then
       pObject.Col = pPiVot
    Else
       pObject.Row = pPiVot
    End If       
    
    iSum = 0
    For iDx = pStart To pEnd
        If pVerHor = "V" Then
           pObject.Row = iDx 
        Else
           pObject.Col = iDx 
        End If
                   
        If Trim(pObject.Text) > ""  Then
           If IsNumeric(unicdbl(pObject.text)) Then
              iSum = iSum + UniCDbl(pObject.text) 
              
           Else
              iOperStatus = False
           End If   
        End If   
        
    Next

    iSumTemp = Replace(iSum    , gClientNum1000, "**"         )
    iSumTemp = Replace(iSumTemp, gClientNumDec ,  "@@"        )
    iSumTemp = Replace(iSumTemp, "@@"          ,   gComNumDec )
    iSum     = Replace(iSumTemp, "**"          ,   gComNum1000)
        
    If iOperStatus = True Then
       If pBool =  True Then
          pObject.Col  = pTargetCol
          pObject.Row  = pTargetRow
          
          pObject.Text = iSum
          
       End If   
    End If   
 
    FncSumSheet  = iSum
    
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

   	pObject.Row = Row
   	pObject.Col = Col

   	If pObject.CellType = SS_CELL_TYPE_FLOAT Then
      If UNICDbl(pObject.Text) < pObject.TypeFloatMin Then
         pObject.Text = UniConvNumPCToCompanyWithoutRound(pObject.TypeFloatMin,0)
      End If
	End If

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
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UNIConvDate = ""
    
    If Trim(pDate) = "" Or IsNull(pDate) Then
       UNIConvDate = gServerBaseDate 
       Exit Function
    End If
    
    If CheckDateFormat(pDate,gServerDateFormat) = True Then                        'If input date format is same as server date format 
       UNIConvDate = pDate
       Exit Function
    End If
    
    pDate = FillLeadingSpaceWithZero(pDate,gDateFormat)

    If CheckDateFormat(pDate,gDateFormat) = False Then                             'If input date format is not same as company date format 
       Exit Function
    End If
    
    Call ExtractDateFrom(pDate,gDateFormat,gComDateType,strYear,strMonth,strDay)  ' From Company Date Type

    UNIConvDate = strYear & gServerDateType & strMonth & gServerDateType & strDay ' To Server Date Type

End Function

'======================================================================================================
Function UNIDateClientFormat(ByVal pDate)
	UNIDateClientFormat =  UNIDateClientFormatSub(pDate,"YMD")
End Function

'======================================================================================================
Function UNIMonthClientFormat(ByVal pDate)
    UNIMonthClientFormat = UNIDateClientFormatSub(pDate,"YM")
End Function

'======================================================================================================
Function UNIDateClientFormatSub(ByVal pDate, ByVal pOption)

    Dim strYear, strMonth, strDay
    Dim iTempDate
	
    On Error Resume Next
    
    UNIDateClientFormatSub = ""
	
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function 
    End If

    pDate = FillLeadingSpaceWithZero(pDate,gAPDateFormat)

    Call ExtractDateFrom(pDate,gAPDateFormat,gAPDateSeperator,strYear,strMonth,strDay)
    
    iTempDate = strYear & gServerDateType & strMonth & gServerDateType & strDay 
    
    If iTempDate <= gServerBaseDate Then 
       Exit Function
    End If

    UNIDateClientFormatSub = MakeDateTo(pOption,gDateFormat,gComDateType,strYear,strMonth,strDay)

End Function

'======================================================================================================
Function UNIConvDateDBToCompany(ByVal pDate, ByVal pDefault)

    UNIConvDateDBToCompany = ""
    
    If Trim(pDate) = "" Or IsNull(pDate) Then
       pDate = pDefault
       If Trim(pDate) = "" Or IsNull(pDate) Then
          Exit Function
       End If
    End If

   UNIConvDateDBToCompany = UNIDateClientFormatSub(pDate,"YMD")
End Function

'======================================================================================================
Function UNIConvDateCompanyToDB(ByVal pDate, ByVal pDefault)

    UNIConvDateCompanyToDB = ""
    
    If Trim(pDate) = "" Or IsNull(pDate) Then
       pDate = pDefault
       If Trim(pDate) = "" Or IsNull(pDate) Then
          Exit Function
       End If
    End If
    
    UNIConvDateCompanyToDB = UNIConvDate(pDate)   	
End Function

'======================================================================================================
Function UNIFormatDate(Byval pDate)
    UNIFormatDate = UniConvLocalToCompanyDateFormat(pDate,"YMD")
End Function

'======================================================================================================
Function UNIFormatMonth(Byval pDate)
    UNIFormatMonth = UniConvLocalToCompanyDateFormat(pDate,"YM")
End Function

'======================================================================================================
Function UniConvLocalToCompanyDateFormat(ByVal pDate, ByVal pOpt)
    Dim strYear, strMonth, strDay
    On Error Resume Next

    UniConvLocalToCompanyDateFormat = ""
    
    If Trim(pDate) = "" Or IsNull(pDate) Then
       Exit Function
    End If		

    pDate = FillLeadingSpaceWithZero(pDate,gClientDateFormat)

    If CheckDateFormat(pDate,gClientDateFormat) = False Then                                 'If input date format is not same as local clientsystem date format 
       Exit Function
    End If

    strYear  =              Year(pDate)
    strMonth = Right("0" & Month(pDate) ,2)
    strDay   = Right("0" &   Day(pDate) ,2)
    
    UniConvLocalToCompanyDateFormat = MakeDateTo(pOpt,gDateFormat,gComDateType,strYear,strMonth,strDay)

End Function

'======================================================================================================
Function UNICDate(ByVal pDate)
    Dim strYear, strMonth, strDay

    On Error Resume Next

    UNICDate = ""

    If Trim(pDate) = "" Or IsNull(pDate) Then
       Exit Function
    End If		
    
    pDate = FillLeadingSpaceWithZero(pDate,gDateFormat)

    If CheckDateFormat(pDate,gDateFormat) = False Then                                    'If input date format is not same as company date format 
       Exit Function
    End If
    
    Call ExtractDateFrom(pDate,gDateFormat  ,gComDateType    ,strYear,strMonth,strDay)
    
    UNICDate = MakeDateTo("YMD",gClientDateFormat,gClientDateSeperator,strYear,strMonth,strDay)
	
End Function

'======================================================================================================
Function UniConvDateToYYYYMMDD(ByVal pDate , ByVal pDateFormat , ByVal pDateSeperator)
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UniConvDateToYYYYMMDD = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If

    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)

    If CheckDateFormat(pDate,pDateFormat) = False Then                                    'If input date format is not same as company date format 
       Exit Function
    End If

    Call ExtractDateFromSuper(pDate,pDateFormat,strYear,strMonth,strDay)

    UniConvDateToYYYYMMDD = strYear & pDateSeperator & strMonth  & pDateSeperator & strDay

End Function

'======================================================================================================
Function UniConvYYYYMMDDToDate(ByVal pDateFormat ,ByVal strYear,ByVal strMonth,ByVal strDay)
    On Error Resume Next
    
    UniConvYYYYMMDDToDate = ""
    
    If IsNull(strYear)   Or Trim(strYear) = "" Then
       Exit Function
    End If
    
    If IsNull(strMonth) Or Trim(strMonth) = "" Then
       Exit Function
    End If
    
    If IsNull(strDay)   Or Trim(strDay)   = "" Then
       Exit Function
    End If

    Select Case pDateFormat
      Case gClientDateFormat : UniConvYYYYMMDDToDate = MakeDateTo("YMD",gClientDateFormat ,gClientDateSeperator,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gAPDateFormat     : UniConvYYYYMMDDToDate = MakeDateTo("YMD",gAPDateFormat     ,gAPDateSeperator    ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gDateFormat       : UniConvYYYYMMDDToDate = MakeDateTo("YMD",gDateFormat       ,gComDateType        ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gDateFormatYYYYMM : UniConvYYYYMMDDToDate = MakeDateTo("YM" ,gDateFormatYYYYMM ,gComDateType        ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gServerDateFormat : UniConvYYYYMMDDToDate = MakeDateTo("YMD",gServerDateFormat ,gServerDateType     ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
	End Select

End Function

'======================================================================================================
Function UniConvDateToYYYYMM(ByVal pDate , ByVal pDateFormat , ByVal pDateSeperator)
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UniConvDateToYYYYMM = UniConvDateToYYYYMMDD(pDate,pDateFormat,pDateSeperator)
    
    If Trim(pDateSeperator) = "" Then
       UniConvDateToYYYYMM = Mid(UniConvDateToYYYYMM,1,6)
    Else
       UniConvDateToYYYYMM = Mid(UniConvDateToYYYYMM,1,6 + Len(pDateSeperator))
    End If

End Function

'======================================================================================================
Function UniConvDateAToB(ByVal pDate , ByVal pFromDateFormat , ByVal pToDateFormat)
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UniConvDateAToB = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If

    pDate = FillLeadingSpaceWithZero(pDate,pFromDateFormat)

    If CheckDateFormat(pDate,pFromDateFormat) = False Then                                    'If input date format is not same as company date format 
       Exit Function
    End If
    
    Call ExtractDateFromSuper(pDate,pFromDateFormat,strYear,strMonth,strDay)

    Select Case pToDateFormat
      Case gClientDateFormat : UniConvDateAToB = MakeDateTo("YMD",gClientDateFormat ,gClientDateSeperator,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gAPDateFormat     : UniConvDateAToB = MakeDateTo("YMD",gAPDateFormat     ,gAPDateSeperator    ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gDateFormat       : UniConvDateAToB = MakeDateTo("YMD",gDateFormat       ,gComDateType        ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gDateFormatYYYYMM : UniConvDateAToB = MakeDateTo("YM" ,gDateFormatYYYYMM ,gComDateType        ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
      Case gServerDateFormat : UniConvDateAToB = MakeDateTo("YMD",gServerDateFormat ,gServerDateType     ,strYear,Right( ("0" & strMonth),2),Right( ("0" & strDay),2))
	End Select

End Function

'======================================================================================================
Function UNIDateAdd(ByVal pInterVal , ByVal pNumber, ByVal pDate, ByVal pDateFormat)
	Dim strYear, strMonth, strDay
	Dim mDate

    On Error Resume Next

    UNIDateAdd = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If
        
    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)

    Call ExtractDateFromSuper(pDate,pDateFormat,strYear,strMonth,strDay)
	
	mDate = strYear & gServerDateType & strMonth & gServerDateType & strDay
	
	mDate = DateAdd(pInterVal,pNumber,mDate) 

    mDate = FillLeadingSpaceWithZero(mDate,gClientDateFormat)
    
    UNIDateAdd = UniConvDateAToB(mDate,gClientDateFormat,pDateFormat)

End Function

'======================================================================================================
Function UNIGetLastDay(ByVal pDate,ByVal pDateFormat)
	Dim strYear, strMonth, strDay
	Dim mDate

    On Error Resume Next
    
    UNIGetLastDay = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If
        
    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)

    Call ExtractDateFromSuper(pDate,pDateFormat,strYear,strMonth,strDay)
	
	If CInt(strMonth) = 12 Then
	   strYear = CInt(strYear) + 1
       strMonth = "01"
	Else
       strMonth = CInt(strMonth) + 1
	End If
	
	strMonth = Right("0" & strMonth ,2)
	
	mDate = strYear & gServerDateType & strMonth & gServerDateType & "01"
	
	mDate = DateAdd("D",-1,mDate) 
	
    mDate = FillLeadingSpaceWithZero(mDate,gClientDateFormat)
    
    If pDateFormat = gDateFormatYYYYMM Then
	   UNIGetLastDay = UniConvDateAToB(mDate,gClientDateFormat,gDateFormat)
	Else   
	   UNIGetLastDay = UniConvDateAToB(mDate,gClientDateFormat,pDateFormat)
	End If   
    
End Function

'======================================================================================================
Function UNIGetFirstDay(ByVal pDate,ByVal pDateFormat)
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UNIGetFirstDay = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If
        
    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)
    
    Call ExtractDateFromSuper(pDate,pDateFormat,strYear,strMonth,strDay)

    If pDateFormat = gDateFormatYYYYMM Then
       UNIGetFirstDay = UniConvYYYYMMDDToDate(gDateFormat ,strYear,strMonth,"01")    
    Else
       UNIGetFirstDay = UniConvYYYYMMDDToDate(pDateFormat ,strYear,strMonth,"01")    
    End If  
End Function

'==============================================================================
Sub ExtractDateFromSuper(ByVal pDate,pDateFormat,strYear,strMonth,strDay)

    If ValidateData(pDate,"SEN") = False Then    '2002/09/28 jinsoo lee
       strYear  = Null
       strMonth = Null
       strDay   = Null
       Exit Sub
    End If
    
    Select Case pDateFormat
          Case gClientDateFormat : Call ExtractDateFrom(pDate,gClientDateFormat , gClientDateSeperator ,strYear,strMonth,strDay)
          Case gAPDateFormat     : Call ExtractDateFrom(pDate,gAPDateFormat     , gAPDateSeperator     ,strYear,strMonth,strDay)
          Case gDateFormat       : Call ExtractDateFrom(pDate,gDateFormat       , gComDateType         ,strYear,strMonth,strDay)
          Case gDateFormatYYYYMM : Call ExtractDateFrom(pDate,gDateFormatYYYYMM , gComDateType         ,strYear,strMonth,strDay)
                                   strDay = "01" 
          Case gServerDateFormat : Call ExtractDateFrom(pDate,gServerDateFormat , gServerDateType      ,strYear,strMonth,strDay)
	End Select

End Sub

'==============================================================================
Function CheckDateFormat(ByVal pDate , ByVal pDateFormat)

    Dim xDate,xDateFormat
    Dim xDateSeperator,cDateFormatArr,cDateArr
    Dim iDx
    Dim strYear,strMonth,strDay
    
    On Error Resume Next
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If
    
    pDate  = UCase(pDate)
    
    xDate = Replace(pDate,"0","X")
    xDate = Replace(xDate,"1","X")
    xDate = Replace(xDate,"2","X")
    xDate = Replace(xDate,"3","X")
    xDate = Replace(xDate,"4","X")
    xDate = Replace(xDate,"5","X")
    xDate = Replace(xDate,"6","X")
    xDate = Replace(xDate,"7","X")
    xDate = Replace(xDate,"8","X")
    xDate = Replace(xDate,"9","X")
    
    
    xDateFormat = Replace(pDateFormat,"Y","X")
    xDateFormat = Replace(xDateFormat,"M","X")
    xDateFormat = Replace(xDateFormat,"D","X")
    
    xDateSeperator = Replace(xDateFormat,"X","")

    xDateSeperator = Mid(xDateSeperator,1,1)
    
    cDateFormatArr = Split(pDateFormat,xDateSeperator)
    
    If Instr(pDate,xDateSeperator) > 0  Then
       cDateArr       = Split(pDate,xDateSeperator)
    Else   
       CheckDateFormat = False
       Exit Function
    End If   
    
    For iDx = 0 To UBound(cDateFormatArr)
        Select Case cDateFormatArr(iDx) 
           Case "YY"   : strYear  = ConvertYYToYYYY(cDateArr(iDx))
           Case "YYYY" : strYear  = cDateArr(iDx)
           Case "MM"   : strMonth = cDateArr(iDx)
           Case "DD"   : strDay   = cDateArr(iDx)
        End Select    
    Next
    
    If Trim(strDay) = "" Then
       strDay = "01"
    End If
    
    If xDate <> xDateFormat Then
       CheckDateFormat = False
    Else
       CheckDateFormat = IsDate(strYear & "-" & strMonth & "-" & strDay)
    End If
End Function

'==============================================================================
Sub ExtractDateFrom(ByVal pDate,pDateFormat,pDateSeperator,strYear,strMonth,strDay)

    strYear = ""
    strDay  = ""
    strDay  = ""
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Sub
    End If

    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)

    pDate = Replace(pDate, pDateSeperator,"")
   
    Select Case Replace(pDateFormat, pDateSeperator,"")
       Case "YYYYMMDD"  : strYear  =                 Mid(pDate,1,4)
                          strMonth =                 Mid(pDate,5,2)
                          strDay   =                 Mid(pDate,7,2)
       Case "YYYYMM"    : strYear  =                 Mid(pDate,1,4)
                          strMonth =                 Mid(pDate,5,2)
       Case "YYMMDD"    : strYear  = ConvertYYToYYYY(Mid(pDate,1,2))
                          strMonth =                 Mid(pDate,3,2) 
                          strDay   =                 Mid(pDate,5,2) 
       Case "YYMM"      : strYear  = ConvertYYToYYYY(Mid(pDate,1,2))
                          strMonth =                 Mid(pDate,3,2) 
       Case "MMDDYYYY"  : strYear  =                 Mid(pDate,5,4)
                          strMonth =                 Mid(pDate,1,2)
                          strDay   =                 Mid(pDate,3,2)
       Case "MMYYYY"    : strYear  =                 Mid(pDate,3,4)
                          strMonth =                 Mid(pDate,1,2)
       Case "MMDDYY"    : strYear  = ConvertYYToYYYY(Mid(pDate,5,2))
                          strMonth =                 Mid(pDate,1,2)
                          strDay   =                 Mid(pDate,3,2)
       Case "MMYY"      : strYear  = ConvertYYToYYYY(Mid(pDate,3,2))
                          strMonth =                 Mid(pDate,1,2)
       Case "DDMMYYYY"  : strYear  =                 Mid(pDate,5,4)
                          strMonth =                 Mid(pDate,3,2)
                          strDay   =                 Mid(pDate,1,2)
       Case "DDMMYY"    : strYear  = ConvertYYToYYYY(Mid(pDate,5,2))
                          strMonth =                 Mid(pDate,3,2)
                          strDay   =                 Mid(pDate,1,2)
       Case "YY"        : strYear  = ConvertYYToYYYY(pDate)
       Case "YYYY"      : strYear  = pDate
    End Select 
    

    If strMonth <> "" Then
       strMonth   = Right("0" & strMonth   , 2)   
    Else   
       strMonth   = "01"
    End If   
    
    If strDay <> "" Then
       strDay   = Right("0" & strDay   , 2)   
    Else   
       strDay   = "01"
    End If   
    
End Sub

'==============================================================================
' Desc : Extract year,month,day from date
'==============================================================================
Function MakeDateTo(pOpt,pDateFormat,pDateSeperator,strYear,strMonth,strDay)

    strMonth = Right("0" & strMonth , 2) 
    strDay   = Right("0" & strDay   , 2) 

    Select Case pOpt
        Case "YMD"
                Select Case Replace(pDateFormat, pDateSeperator,"")
                    Case "YYYYMMDD"  :  MakeDateTo = strYear                  & pDateSeperator & strMonth                 & pDateSeperator & strDay
                    Case "YYMMDD"    :  MakeDateTo = Mid(strYear,3,2)         & pDateSeperator & strMonth                 & pDateSeperator & strDay
                    Case "MMDDYYYY"  :  MakeDateTo = strMonth                 & pDateSeperator & strDay                   & pDateSeperator & strYear
                    Case "MMDDYY"    :  MakeDateTo = strMonth                 & pDateSeperator & strDay                   & pDateSeperator & Mid(strYear,3,2)
                    Case "DDMMYYYY"  :  MakeDateTo = strDay                   & pDateSeperator & strMonth                 & pDateSeperator & strYear
                    Case "DDMMYY"    :  MakeDateTo = strDay                   & pDateSeperator & strMonth                 & pDateSeperator & Mid(strYear,3,2)
                End Select     
		Case "YM"
                Select Case Replace(pDateFormat, pDateSeperator,"")
                    Case "YYYYMMDD"  :  MakeDateTo = strYear                  & pDateSeperator & strMonth
                    Case "YYMMDD"    :  MakeDateTo = Mid(strYear,3,2)         & pDateSeperator & strMonth
                    Case "MMDDYYYY"  :  MakeDateTo = strMonth                 & pDateSeperator & strYear
                    Case "MMDDYY"    :  MakeDateTo = strMonth                 & pDateSeperator & Mid(strYear,3,2)
                    Case "DDMMYYYY"  :  MakeDateTo = strMonth                 & pDateSeperator & strYear
                    Case "DDMMYY"    :  MakeDateTo = strMonth                 & pDateSeperator & Mid(strYear,3,2)
                    Case "YYYYMM"    :  MakeDateTo = strYear                  & pDateSeperator & strMonth
                    Case "YYMM"      :  MakeDateTo = Mid(strYear,3,2)         & pDateSeperator & strMonth
                    Case "MMYY"      :  MakeDateTo = strMonth                 & pDateSeperator & Mid(strYear,3,2)
                    Case "MMYYYY"    :  MakeDateTo = strMonth                 & pDateSeperator & strYear
                End Select      
		Case "MD"
                Select Case Replace(pDateFormat, pDateSeperator,"")
                    Case "DDMMYYYY", "DDMMYY" : 
                                        MakeDateTo = strDay                   & pDateSeperator & strMonth
                    Case Else : 
                                        MakeDateTo = strMonth                 & pDateSeperator & strDay                   
                End Select      
                
    End Select

End Function

'==============================================================================
Function FillLeadingSpaceWithZero(ByVal pDate,pDateFormat)
    Dim tmpArrDate
    Dim tmpArrDateFormat
    Dim pDateSeperator
    Dim iLoop
    
    On Error Resume Next
    
    FillLeadingSpaceWithZero = ""

    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If

    Select Case pDateFormat
        Case gClientDateFormat : pDateSeperator = gClientDateSeperator
        Case gAPDateFormat     : pDateSeperator = gAPDateSeperator
        Case gDateFormat       : pDateSeperator = gComDateType
        Case gDateFormatYYYYMM : pDateSeperator = gComDateType
        Case gServerDateFormat : pDateSeperator = gServerDateType
	End Select
    
    tmpArrDateFormat = Split(pDateFormat,pDateSeperator)
    tmpArrDate       = Split(pDate      ,pDateSeperator)
    
    For iLoop = 0 To UBound(tmpArrDateFormat)
       Select Case tmpArrDateFormat(iLoop)
          Case "YYYY"  : tmpArrDate(iLoop) = tmpArrDate(iLoop) 
          Case "YY"    : tmpArrDate(iLoop) = tmpArrDate(iLoop) 
          Case "MM"    : tmpArrDate(iLoop) = Right("0" & tmpArrDate(iLoop),2)
          Case "DD"    : tmpArrDate(iLoop) = Right("0" & tmpArrDate(iLoop),2)
       End Select    

       If iLoop = 0 Then 
          FillLeadingSpaceWithZero  = FillLeadingSpaceWithZero  &  tmpArrDate(iLoop)
       Else
          FillLeadingSpaceWithZero  = FillLeadingSpaceWithZero  & pDateSeperator & tmpArrDate(iLoop)
       End If   
    Next   

End Function

'==============================================================================
Function ConvertYYToYYYY(pYY)

    ConvertYYToYYYY = ""
    
    If IsNull(pYY) Or Trim(pYY) = "" Then
       Exit Function
    End If
    
    If CDbl(pYY) > 50 Then
        ConvertYYToYYYY =  "19" & pYY
    Else
        ConvertYYToYYYY =  "20" & pYY
    End If
    
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

	On Error Resume Next
	
    If ValidateData(pNum,"SEN") = False Then
       UNICCur = 0
       Exit Function       
    End If
	
	pNum = CStr(pNum)

	pNum = Replace(pNum, " "        , "")
	pNum = Replace(pNum, gComNum1000, "")
    pNum = Replace(pNum, gComNumDec , gClientNumDec) 

	UNICCur = CCur(pNum)    

End Function

'==============================================================================
Function UNIConvNum(ByVal pNum, ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       UNIConvNum = pDefault
       Exit Function       
    End If
  	    	    
    pNum       = Replace(pNum      , " "  , "")
    pNum       = Replace(pNum, gComNum1000, "")
    
    UNIConvNum = Replace(pNum, gComNumDec , ".")

End Function

'======================================================================================================
Function UNICDbl(ByVal pNum)

    If ValidateData(pNum,"SEN") = False Then
       UNICDbl = 0
       Exit Function       
    End If
    
    pNum    = CStr(pNum)
		
    pNum    = Replace(pNum, " "        , "" )
    pNum    = Replace(pNum, gComNum1000, "" )
    pNum    = Replace(pNum, gComNumDec , gClientNumDec)

    UNICDbl = CDbl(pNum)

End Function

'======================================================================================================
Function UniConvNumPCToCompanyWithoutRound(ByVal pNum,ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault
       Exit Function       
    End If
	
    UniConvNumPCToCompanyWithoutRound = uniConvNumAToB(pNum,gClientNum1000,gClientNumDec, gComNum1000, gComNumDec,True,"X","X") 

End Function

'======================================================================================================
Function UNIFormatNumber(ByVal pNum, ByVal pDecPoint, ByVal pFormatType, ByVal pNegativeNum,ByVal pRndPolicy, ByVal pRndUnit)

    Dim retVal 
    Dim iTmpNum
    Dim iTmpNumArr
    Dim iTmpNumInt,iTmpNumDec

    If ValidateData(pNum,"SEN") = False Then
       If pNegativeNum = 1 Then
          UNIFormatNumber = ""
          Exit Function
       Else
          UNIFormatNumber = Replace(FormatNumber(0, pDecPoint, pFormatType), gClientNumDec , gComNumDec)
          Exit Function
       End If
    End If
	
    If CDbl(pNum) = 0 Then
       UNIFormatNumber = Replace(FormatNumber(0, pDecPoint, pFormatType), gClientNumDec , gComNumDec)
       Exit Function
    End If
    
    If IsNumeric(pDecPoint) Then
       pDecPoint = CInt(pDecPoint)
    Else
       pDecPoint = 0
    End If   

    pNum = Replace(CStr(pNum) , " "            , "")
    pNum = Replace(     pNum  , gClientNum1000 , "")
    
    pNum = MakeExpNumToStrNum(pNum)
    
    pNum = FncRoundData(pNum,pDecPoint,pRndPolicy,pRndUnit,gClientNumDec)
    
    UNIFormatNumber = uniConvNumAToB(pNum,gClientNum1000,gClientNumDec, gComNum1000, gComNumDec,True,"X","X")        
    
End Function

'========================================================================================
Function FncRoundData(ByVal pNum, ByVal pDecPoint, ByVal pRndPolicy, ByVal pRndUnit, ByVal pNumDec)
   Dim iNumARR(30)
   Dim iNumSTR
   Dim ii, jj
   Dim iBit
   Dim iNumDFindPlace
   Dim iCountAfterDec
   Dim iMeetDPoint
   Dim iPositive
   Dim iDataMaxLength
   
   If CDbl(pNum) > 0 Then
      iPositive = True
   Else
      iPositive = False
   End If

   If IsNumeric(pDecPoint) Then
      pDecPoint = CInt(pDecPoint)
   Else
      pDecPoint = 0
   End If
    
   If InStr(CStr(pNum), pNumDec) = 0 Then           'If it is Integer
      If pDecPoint > 0 Then
         FncRoundData = pNum & pNumDec & Mid("000000", 1, pDecPoint)
      Else
         FncRoundData = pNum
      End If
      Exit Function
   End If

   jj = 0
   iNumARR(jj) = 0
   iCountAfterDec = 0
   iMeetDPoint = False
   iNumDFindPlace = 0
   
   For ii = 1 To Len(pNum)
       If iMeetDPoint = True Then
          iCountAfterDec = iCountAfterDec + 1
          If iCountAfterDec > pDecPoint + 1 Then
              Exit For
          End If
       End If
       iBit = Mid(pNum, ii, 1)
       If iBit = pNumDec Then
          iMeetDPoint = True
          iNumDFindPlace = jj + 1

       ElseIf IsNumeric(iBit) = True Then
          jj = jj + 1
          iNumARR(jj) = iBit
       End If
   Next
   
   If iCountAfterDec < pDecPoint Then
      FncRoundData = pNum & Mid("000000", 1, pDecPoint - iCountAfterDec)
      Exit Function
   End If
   
   iDataMaxLength = iNumDFindPlace + pDecPoint
   Select Case pRndPolicy
         Case "1":                                                  ' 올림 
                  If CInt(iNumARR(iDataMaxLength)) > 0 Then
                     iNumARR(iDataMaxLength - 1) = iNumARR(iDataMaxLength - 1) + 1
                  End If
         Case "2"                                                   ' 내림 
         Case "3"                                                   ' 반올림 
                  If CInt(iNumARR(iNumDFindPlace + pDecPoint)) > 4 Then
                     iNumARR(iDataMaxLength - 1) = iNumARR(iDataMaxLength - 1) + 1
                  End If
   End Select
   
   iNumARR(iNumDFindPlace + pDecPoint) = 0
   
   For ii = iDataMaxLength To 1 Step -1
       If iNumARR(ii) > 9 Then
          iNumARR(ii) = iNumARR(ii) - 10
          iNumARR(ii - 1) = iNumARR(ii - 1) + 1
       End If
   Next
   
   iNumSTR = ""
   
   If iNumARR(0) > 0 Then
       iNumSTR = iNumARR(0) & iNumSTR
   End If
   
   For ii = 1 To iDataMaxLength - 1
       If iNumDFindPlace = ii Then
          iNumSTR = iNumSTR & pNumDec
       End If
       iNumSTR = iNumSTR & iNumARR(ii)
   Next
   
   If iPositive = False Then
      FncRoundData = "-" & iNumSTR
   Else
      FncRoundData = iNumSTR
   End If
End Function

'==============================================================================
Function MarkNum1000SEP(ByVal pNum,ByVal pALTNum,ByVal p1000SEP)  

    Dim iData
    
    MarkNum1000SEP = ""
    
    If ValidateData(pNum,"SENF") = False Then
       MarkNum1000SEP = pALTNum       
       Exit Function
    End If
    
    If Len(Trim(pNum)) < 4 Then
       MarkNum1000SEP = pNum  
       Exit Function
    End If    

    pNum = StrReverse(pNum)
    
    Do While 1
  
       iData = iData & Mid(pNum,1,3)
       pNum  = Mid(pNum,4)

       If Trim(pNum) = "" Then
          Exit Do              ' Exit loop.
       End If
       iData       = iData & p1000SEP 
    Loop
    
    iData          = StrReverse(iData)

    If Mid(iData,1,2) = "-," Then
       MarkNum1000SEP = "-" & Mid(iData,3)
       Exit Function
    End If
    
    If Mid(iData,1,2) = "-." Then
       MarkNum1000SEP = "-" & Mid(iData,3)
       Exit Function
    End If
    
    MarkNum1000SEP = iData
      
End Function   

'=========================================================================================================================
Function uniConvNumAToB(ByVal pNum, ByVal pNum1000From, ByVal pNumDecFrom, ByVal pNum1000To, ByVal pNumDecTo, ByVal p1000SEP,ByVal pOpt1,ByVal pOpt2)
    Dim iTmpNumInt
    Dim iTmpNumDec
    Dim iTmpNumArr

    pNum = Replace(pNum, " "         , "")
    pNum = Replace(pNum, pNum1000From, "")

    If InStr(pNum, pNumDecFrom) > 0 Then
       iTmpNumArr = Split(pNum, pNumDecFrom)
       iTmpNumInt = iTmpNumArr(0)
       iTmpNumDec = iTmpNumArr(1)
    Else
       iTmpNumInt = pNum
       iTmpNumDec = ""
    End If
    
    If p1000SEP = True Then
       If InStr(iTmpNumInt, pNum1000From) > 0 Then
          iTmpNumInt = Replace(iTmpNumInt, pNum1000From, pNum1000To)
       Else
          iTmpNumInt = MarkNum1000SEP(iTmpNumInt, "0", pNum1000To)
       End If
    End If
    
    If iTmpNumDec = "" Then
       uniConvNumAToB = iTmpNumInt
    Else
       uniConvNumAToB = iTmpNumInt & pNumDecTo & iTmpNumDec
    End If

End Function

'=========================================================================================================================
Function MakeExpNumToStrNum(ByVal pNum)
   Dim tmpExpArr
   Dim tmpNumberArr
   Dim tmpZeroString
   Dim iPositive
   Dim tmpData
   
   If CDbl(pNum) > 0 Then
      iPositive = True
   Else
      iPositive = False
   End If
   
   tmpZeroString = "0000000000000000000000000"
   
   If InStr(pNum, "E+") Then
      tmpExpArr = Split(pNum, "E+")
      If InStr(tmpExpArr(0), ".") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ".")
         tmpData = tmpNumberArr(0) & Left(tmpNumberArr(1) & tmpZeroString, tmpExpArr(1))
      ElseIf InStr(tmpExpArr(0), ",") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ",")
         tmpData = tmpNumberArr(0) & Left(tmpNumberArr(1) & tmpZeroString, tmpExpArr(1))
      Else
         tmpData = tmpExpArr(0) & Left(tmpZeroString, tmpExpArr(1))
      End If
   ElseIf InStr(pNum, "E-") Then
      tmpExpArr = Split(pNum, "E-")
      If InStr(tmpExpArr(0), ".") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ".")
         tmpData = "0." & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpNumberArr(0) & tmpNumberArr(1)
      ElseIf InStr(tmpExpArr(0), ",") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ",")
         tmpData = "0," & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpNumberArr(0) & tmpNumberArr(1)
      Else
         tmpData = "0." & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpExpArr(0)
      End If
   Else
      tmpData = pNum
   End If
   
   If InStr(tmpData,"-") Then
      MakeExpNumToStrNum = "-" & Replace(tmpData,"-","")
   Else   
      MakeExpNumToStrNum = tmpData
   End If

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
          ggStrIntegeralPart = ggStrIntegeralPart & gColSep 
       Next   
    End If
    
    If Trim(ggStrDeciPointPart) = "" Then 
       For iDx = 0 To 13 
          ggStrDeciPointPart = ggStrDeciPointPart & gColSep 
       Next   
    End If

    ggStrIntegeralPart =  Split(ggStrIntegeralPart,gColSep)
    ggStrDeciPointPart =  Split(ggStrDeciPointPart,gColSep)
    
    If Trim(iIntegeral) = "" Then
       iIntegeral = 15 - CInt(iDec)
    End If
    
    ggStrIntegeralPart(iiPos) = CStr(iIntegeral)
    ggStrDeciPointPart(iiPos) = CStr(iDec)
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    For iDx = 0 To 13
        sBuffer1 =  sBuffer1 & ggStrIntegeralPart(iDx) & gColSep
        sBuffer2 =  sBuffer2 & ggStrDeciPointPart(iDx) & gColSep
    Next
    
    ggStrIntegeralPart = sBuffer1
    ggStrDeciPointPart = sBuffer2

End Sub

'===============================================================================
Sub AppendNumberRange(ByVal iPos,ByVal iMin,ByVal iMax)
    Dim iDx
    Dim iiPos
    Dim sBuffer1
    Dim sBuffer2

    iiPos = CInt(iPos)
    
    If Trim(ggStrMinPart) = "" Then 
       For iDx = 0 To 10 
          ggStrMinPart = ggStrMinPart & gColSep 
       Next   
    End If    '

    If Trim(ggStrMaxPart) = "" Then 
       For iDx = 0 To 10 
          ggStrMaxPart = ggStrMaxPart & gColSep 
       Next   
    End If       

    ggStrMinPart =  Split(ggStrMinPart,gColSep)
    ggStrMaxPart =  Split(ggStrMaxPart,gColSep)
    
    ggStrMinPart(iiPos) = CStr(iMin)
    ggStrMaxPart(iiPos) = CStr(iMax)
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    For iDx = 0 To 9
        sBuffer1 =  sBuffer1 & ggStrMinPart(iDx) & gColSep
        sBuffer2 =  sBuffer2 & ggStrMaxPart(iDx) & gColSep
    Next
    ggStrMinPart = sBuffer1
    ggStrMaxPart = sBuffer2

End Sub

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
Function ChkField(pDoc, ByVal pStrGrp)
    On Error Resume Next
    
    Dim i, intDivCnt, intTagNum
    Dim strTagName, strRequired
    Dim iRet
    
    iRequired = UCase(UCN_REQUIRED)
            
    intDivCnt = 0
    ChkField = False
    
    For i = 0 To pDoc.All.Length - 1
    
        iRet = ChkFieldByCell(pDoc.All(i), pStrGrp, intDivCnt)
        If iRet = False Then
           Exit Function
        End If
        
    Next
    
    ChkField = ChkFieldLength(pDoc, pStrGrp)
    
End Function
  
'========================================================================
Function ChkFieldByCell(TmpObject, ByVal pStrGrp, intDivCnt)
       Dim strTagName
       Dim intTagNum
       Dim strRequired
       Dim iRet
       
        On Error Resume Next
       
        strTagName  = ""
        intTagNum   = 0
        strRequired = ""
        
        ChkFieldByCell = False

        strTagName = UCase(TmpObject.tagName)
        
        If strTagName <> Empty Then
           If strTagName = "DIV" Then
              intDivCnt = intDivCnt + 1
           End If
        End If
        
        If Not( strTagName = "INPUT" Or strTagName = "TEXTAREA" Or strTagName = "SELECT" Or strTagName = "OBJECT") Then
           ChkFieldByCell = True
           Exit Function
        End If        
                
        If UCase(TypeName(TmpObject.getAttribute("tag"))) = "NULL" Then
           ChkFieldByCell = True
           Exit Function
        End If
                
        intTagNum = Mid(TmpObject.getAttribute("tag"), 1, 1)
        strRequired = UCase(TmpObject.className)
                
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If (intTagNum = pStrGrp Or pStrGrp = "A") And strRequired = UCase(UCN_REQUIRED) Then
                Select Case strTagName
                    Case "INPUT", "TEXTAREA", "SELECT"     
                        If Len(Trim(TmpObject.Value)) = 0 Then
                            If intTagNum = "1" Then
                                iRet = DisplayMsgBox("970029", "X", TmpObject.alt, "x")
                            Else
                                iRet = DisplayMsgBox("970021", "X", TmpObject.alt, "x")
                            End If
                            Call ChangeTabs2(Document,intDivCnt)
                            TmpObject.focus
                            Set gActiveElement = Document.activeElement
                            Exit Function
                        End If
                        
                    Case "OBJECT"
                        If TmpObject.Title = "FPDATETIME" Or TmpObject.Title = "FPDOUBLESINGLE" Then
                            If Len(Trim(TmpObject.Text)) = 0 Then
                                If intTagNum = "1" Then
                                   iRet = DisplayMsgBox("970029", "X", TmpObject.alt, "x")
                                Else
                                   iRet = DisplayMsgBox("970021", "X", TmpObject.alt, "x")
                                End If
                                
                            Call ChangeTabs2(Document,intDivCnt)
                                Call SetFocusToDocument("M")
                                TmpObject.focus
                                
                                Set gActiveElement = Document.activeElement
                                Exit Function
                            End If
                        End If
                        
                End Select
                
            End If
            
        End If

        ChkFieldByCell = True
    
End Function
  
'========================================================================
Function ChkFieldLength(pDoc, ByVal pStrGrp)
    On Error Resume Next
    
    Dim i, intDivCnt
    Dim iRet

    intDivCnt = 0

    ChkFieldLength = False
    
    For i = 0 To pDoc.All.Length - 1
        iRet = ChkFieldLengthByCell(pDoc.All(i),pStrGrp,intDivCnt)
        If iRet = False Then
           Exit Function
        End If        
    Next
    
    ChkFieldLength = True
    
End Function  

'========================================================================
Function ChkFieldLengthByCell(pTempDoc, ByVal pStrGrp, intDivCnt)
        Dim strTagName
        Dim intTagNum
        Dim strRequired
        Dim iMaxLen, iRet
        
        On Error Resume Next
        
        ChkFieldLengthByCell = Fasle
     
        strTagName  = ""
        intTagNum   = 0
        strRequired = ""
        
        strTagName = UCase(pTempDoc.tagName)
        
        If strTagName <> Empty Then
           If strTagName = "DIV" Then
              intDivCnt = intDivCnt + 1
              ChkFieldLengthByCell = True
              Exit Function
           End If
        End If
                
        If strTagName <> "INPUT" Then
           ChkFieldLengthByCell = True
           Exit Function
        End If        
        
        If UCase(TypeName(pTempDoc.getAttribute("tag"))) = "NULL" Then
           ChkFieldLengthByCell = True
           Exit Function
        End If
                
        intTagNum = Mid(pTempDoc.Tag, 1, 1)
        strRequired = UCase(pTempDoc.className)
        
       If Not (intTagNum = pStrGrp Or pStrGrp = "A") Then
          ChkFieldLengthByCell = True
          Exit Function
       End If   
        
        
        If Err.Number = 0 Then
           If UCase(pTempDoc.Type) = "TEXT" Then
              iMaxLen = CDbl(pTempDoc.MaxLength)
              If strRequired <> UCase(UCN_PROTECTED) Then
                 If CmpCharLength(Trim(pTempDoc.Value), iMaxLen) = False Then
                    iRet = DisplayMsgBox("900028", "X", pTempDoc.alt, "X")
                    Call ChangeTabs2(Document,intDivCnt)
                    pTempDoc.focus
                    Set gActiveElement = Document.activeElement
                    Exit Function
                  End If
              End If
           End If
            
        End If
        
        ChkFieldLengthByCell = True

End Function

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
Function CheckRunningBizProcess()

	CheckRunningBizProcess = True

	If window.document.all("MousePT").style.visibility = "visible" Then 
	   Exit Function
	End If   

	CheckRunningBizProcess = False

End Function

'===============================================================================
Function CompareDateByFormat(pFromDt, pToDt,pFromDtAlt,pToDtAlt,ByVal pMsgCD,ByVal pDateFormat,ByVal pDateSeperator,pBool)

    Dim strYear1,strMonth1,strDay1,strFullDay1
    Dim strYear2,strMonth2,strDay2,strFullDay2

	CompareDateByFormat = False
    
    Call ExtractDateFrom(pFromDt,pDateFormat,pDateSeperator,strYear1,strMonth1,strDay1)
    Call ExtractDateFrom(pToDt  ,pDateFormat,pDateSeperator,strYear2,strMonth2,strDay2)
      
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
Function CompareDateByFormat2(pFromDt, pToDt,pFromDtAlt,pToDtAlt,ByVal pMsgCD,ByVal pDateFormat,ByVal pDateSeperator,pBool)

    Dim strYear1,strMonth1,strDay1,strFullDay1
    Dim strYear2,strMonth2,strDay2,strFullDay2

	CompareDateByFormat2 = False
    
    Call ExtractDateFrom(pFromDt,pDateFormat,pDateSeperator,strYear1,strMonth1,strDay1)
    Call ExtractDateFrom(pToDt  ,pDateFormat,pDateSeperator,strYear2,strMonth2,strDay2)
      
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
Sub ReFormatSpreadCellByCellByCurrency(pObject,ByVal pStartRow,ByVal pEndRow,ByVal pCurrencyCol,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, ByVal Dummy1, ByVal Dummy2)
    Dim ii
    Dim iData
    Dim iCurrency
    Dim iDecimalPlaceAlignOpt
    Dim iDx
    Dim iArrDec, iDefaultDec, iDataType
    
    If UCase(pFormType) = "Q" Then
        iDecimalPlaceAlignOpt = gQMDPAlignOpt
    Else
        iDecimalPlaceAlignOpt = gIMDPAlignOpt
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
    iArrDec = Split(ggStrDeciPointPart,gColsep)
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
            iData = UNICdbl(pObject.Text)
            pObject.Text = UNIFormatNumber(iData,gBDecimals(iDx) , -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx))
            If iDecimalPlaceAlignOpt = "1" Then
                pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
            Else
                pObject.TypeFloatDecimalPlaces = iDefaultDec
            End If
        End If
    Next
End Sub

'=================================================================================================================
Sub ReFormatSpreadCellByCellByCurrency2(pObject,ByVal pStartRow,ByVal pEndRow,ByVal pCurrency,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, ByVal Dummy1, ByVal Dummy2)
    Dim ii
    Dim iData
    Dim iDecimalPlaceAlignOpt
    Dim iDx
    Dim iArrDec, iDefaultDec, iDataType    
    
'    If UCase(pFormType) = "Q" Then
'        iDecimalPlaceAlignOpt = gQMDPAlignOpt
'    Else
'        iDecimalPlaceAlignOpt = gIMDPAlignOpt
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
    iArrDec = Split(ggStrDeciPointPart,gColsep)
    iDefaultDec = iArrDec(iDataType + 8)
    iDx = FindIndexOfCurrency(pCurrency,iDataType)
    For ii = pStartRow to pEndRow
        pObject.Row = ii
        pObject.Col = pTargetCol
        If iDx = -1 Then 
            pObject.TypeFloatDecimalPlaces = iDefaultDec        
        Else
            iData = UNICdbl(pObject.Text)
            pObject.Text = UNIFormatNumber(iData,gBDecimals(iDx) , -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx))
 '           If iDecimalPlaceAlignOpt = "1" Then
                pObject.TypeFloatDecimalPlaces = gBDecimals(iDx)
 '           Else
 '               pObject.TypeFloatDecimalPlaces = iDefaultDec
 '           End If
        End If
    Next
End Sub

'=================================================================================================================
Sub EditModeCheck(pObject,ByVal pRow,ByVal pCurrencyCol,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, Byval pMode, ByVal Dummy1, ByVal Dummy2)
    Dim iCurrency
    Dim iArrDec
    Dim iDecimalPlaceAlignOpt
    Dim iDx, iDataType
    
    If UCase(pFormType) = "Q" Then
        iDecimalPlaceAlignOpt = gQMDPAlignOpt
    Else
        iDecimalPlaceAlignOpt = gIMDPAlignOpt
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
        iArrDec = Split(ggStrDeciPointPart,gColsep)
        pObject.Col = pTargetCol
        pObject.TypeFloatDecimalPlaces = iArrDec(iDataType + 8)
    End If
End Sub

'=================================================================================================================
Sub EditModeCheck2(pObject,ByVal pRow,ByVal pCurrency,ByVal pTargetCol,ByVal pDataType ,ByVal pFormType, Byval pMode, ByVal Dummy1, ByVal Dummy2)
    Dim iArrDec
    Dim iDecimalPlaceAlignOpt
    Dim iDx, iDataType
    
'    If UCase(pFormType) = "Q" Then
'        iDecimalPlaceAlignOpt = gQMDPAlignOpt
'    Else
'        iDecimalPlaceAlignOpt = gIMDPAlignOpt
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
        iArrDec = Split(ggStrDeciPointPart,gColsep)
        pObject.Col = pTargetCol
        pObject.TypeFloatDecimalPlaces = iArrDec(iDataType + 8)
    End If
End Sub

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
        iData = Fix(CStr(UNICDbl(pObject.Text) * iTemp)) / iTemp
        pObject.Text = UniConvNumPCToCompanyWithoutRound(iData,"0")
    End If
End Sub

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
        iData = Fix(CStr(UNICDbl(pObject.Text) * iTemp)) / iTemp
        pObject.Text = UniConvNumPCToCompanyWithoutRound(iData,"0")
    End If
End Sub


'======================================================================================================
' Function Name : DisplayMsgBox
' Function Desc : 메세지 테이블을 검색하여 결과를 리턴한다.
'======================================================================================================
Function DisplayMsgBox(ByVal pMsgId,ByVal pBtnKind,ByVal pMsg1,ByVal pMsg2)
    Dim iCount
    Dim iRet
    Dim iRet1
    Dim iRet2
    Dim iRet3
    
    ReDim iRet2(4)

    If UCase(Trim(CStr(pBtnKind))) = "X" Or UCase(Trim(CStr(pBtnKind))) = "" Then
       pBtnKind = "0"
    End If

    If UCase(Trim(CStr(pMsg1))) = "X" Then
       pMsg1 = ""
    End If
    
    If UCase(Trim(CStr(pMsg2))) = "X" Then
       pMsg2 = ""
    End If

    If Len(Trim(pMsgId)) = 6 Then
      
       If  FetchBMessage(Cstr(pMsgId),iRet ) = True Then
           iRet2 = Split(iRet,Chr(12))
           If UCase(iRet2(0)) = "Y" Then
              iCount = CountStrings(iRet2(1), "%")
              If iCount > 0 Then
                 iRet2(1) = MessageSplit(iCount,iRet2(1), pMsg1, pMsg2)
              End If
           End IF
       End IF
       
    ElseIf Len(Trim(pMsgId)) > 6 Then
    
       If M990013 = "" Then
          If FetchBMessage("990013",iRet) = True Then
             iRet3 = split(iRet,chr(12))       
             M990013 = iRet3(1)
          End If   
       End If    
       iRet2(0) = "X"
       iRet2(1) = M990013 & vbcrlf & vbcrlf & "Message code : " & pMsgId
    
    ElseIf Trim(pMsg1) > "" Then
       iRet2(0) = "X"
       iRet2(1) = pMsg1
    Else
       If M990012 = "" Then
          If FetchBMessage("990012",iRet) = True Then
             iRet3 = split(iRet,chr(12))       
             M990012 = iRet3(1)
          End If   
       End If   
       iRet2(0) = "X"
       iRet2(1) = M990012 & vbcrlf & vbcrlf & "Message code : " & pMsgId
    End If   
    
    If iRet2(0) = "X" Then
       iRet2(2) = "1"             'Default value set
       pBtnKind = "0"
    End If
    
    If pBtnKind = "0" Then
        Select Case iRet2(2)
            Case "1"   ' Information
                   DisplayMsgBox = MsgBox(iRet2(1), vbInformation, gLogoName & "-[Information]")
            Case "2"   ' Warning
                   DisplayMsgBox = MsgBox(iRet2(1), vbExclamation, gLogoName & "-[Warning]")
            Case "3"   ' Error
                   DisplayMsgBox = MsgBox(iRet2(1), vbCritical   , gLogoName & "-[Error]")
            Case "4"   ' Fatal   
                   DisplayMsgBox = MsgBox(iRet2(1), vbCritical   , gLogoName & "-[Fatal]")
            Case Else       
            
                   If M990014 = "" Then
                      If FetchBMessage("990014",iRet) = True Then
                         iRet3 = split(iRet,chr(12))       
                         M990014 = iRet3(1)
                      End If   
                   End If
                      
                   If Trim(M990014) <> "" Then
                      DisplayMsgBox = MsgBox(M990014 & vbCrLf & vbCrLf & "Error Level : " & iRet2(2) , vbInformation   , gLogoName & "-[Information]")
                   End If   
                   
        End Select

    Else
        Select Case CInt(pBtnKind)
            Case 33   ' Ok, Cancel
                   DisplayMsgBox = MsgBox(iRet2(1), vbOKCancel    + vbQuestion, gLogoName)
            Case 35   ' Yes,No,Cancel
                   DisplayMsgBox = MsgBox(iRet2(1), vbYesNoCancel + vbQuestion, gLogoName)
            Case 36   ' Yes,No
                   DisplayMsgBox = MsgBox(iRet2(1), vbYesNo       + vbQuestion, gLogoName)
            Case 64   ' Information
                   DisplayMsgBox = MsgBox(iRet2(1), vbInformation + vbQuestion, gLogoName)

            Case Else

                   If M990015 = "" Then
                      If FetchBMessage("990015",iRet) = True Then
                         iRet3 = split(iRet,chr(12))       
                         M990015 = iRet3(1)
                      End If   
                   End If
                      
                   If Trim(M990015) <> "" Then
                      DisplayMsgBox = MsgBox(M990015 & vbCrLf & vbCrLf & "Button Code : " & pBtnKind, vbInformation, gLogoName & "-[Information]")
                   End If   

        End Select
    End If       
    
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
      
    If gRdsUse = "T" Then
       Set pRDSCom = ADS.CreateObject("PuniRDSCom.CuniRDSCom",gServerIP  )
       prData = pRDSCom.GetMessageData(pCode, gEnvInf)
       Set pRDSCom = Nothing
    Else
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
    End If
    
    If Err.number <> 0 Then
        Exit Function
    End If
    
    FetchBMessage = True
    
End Function



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
Sub SetFocusToDocument(pOpt)
   Select Case pOpt
     Case "M"
         top.Window.Parent.Frames(1).Focus	                 
     Case "P" 
          window.focus
   End Select      
End Sub

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
Sub SetCombo2(pCombo, ByVal pCodeArr, ByVal pNameArr,pSeperator)

    Dim iDx

    pCodeArr = Split(pCodeArr,pSeperator)
    pNameArr = Split(pNameArr,pSeperator)
    
    For iDx = 0 To UBound(pCodeArr) - 1
        Call SetCombo(pCombo,pCodeArr(iDx), pNameArr(iDx))
    Next

End Sub

'======================================================================================================
Sub SetSpreadFloat(ByVal iCol ,ByVal Header ,ByVal dColWidth ,ByVal HAlign ,ByVal iFlag )
    ggoSpread.SSSetFloat iCol,Header,dColWidth,CStr(iFlag),ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec,HAlign
End Sub

'===============================================================================================
Function UNIFormatNumberByCurrecny(ByVal pNum,ByVal pCurrency,ByVal pDataType)
    UNIFormatNumberByCurrecny = UNIConvNumPCToCompanyByCurrency(pNum, pCurrency, pDataType,"X", "X")    
End Function

'===============================================================================================
Function uniFormatNumberByTax(ByVal pNum,ByVal pCurrency,ByVal pDataType)

    If ValidateData(pDataType,"SEN") = False Then
       pDataType = ggAmtOfMoneyNo
	End If
	
    uniFormatNumberByTax = UNIConvNumPCToCompanyByCurrency(pNum, pCurrency, pDataType,gTaxRndPolicyNo, "X")

End function

'===============================================================================================
Function UNIConvNumPCToCompanyByCurrency(ByVal pNum,ByVal pCurrency,ByVal pDataType,ByVal pOpt1,ByVal pOpt2)

    Dim iDx
    Dim iRet
	
    UNIConvNumPCToCompanyByCurrency = ""
   
    iDx = FindIndexOfCurrency(pCurrency,pDataType)
	
    If CInt(iDx) < 0 Then 
       iDx = FindIndexOfCurrency(gCurrency,pDataType)

       If CInt(iDx) < 0 Then 
          iRet = MsgBox ("화폐별 포맷정보를 찾을 수가 없습니다." ,vbExclamation,gLogoName)  '2002/08/13 lee jinsoo
          UNIConvNumPCToCompanyByCurrency = UniConvNumPCToCompanyWithoutRound(pNum,"")       '2002/08/13 lee jinsoo
          Exit Function
        End If   
    End If
	
    Select Case pOpt1
       Case gTaxRndPolicyNo   : UNIConvNumPCToCompanyByCurrency = UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gTaxRndPolicy        ,gBRoundingUnit(iDx))
       Case gLocRndPolicyNo   : 
                             If gBConfMinorCD = "1" Then
                                UNIConvNumPCToCompanyByCurrency = UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx)) 
                             Else
                                UNIConvNumPCToCompanyByCurrency = UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gLocRndPolicy        ,gBRoundingUnit(iDx))
                             End If   
       Case Else              : UNIConvNumPCToCompanyByCurrency = UNIFormatNumber(pNum, gBDecimals(iDx), -2, 0,gBRoundingPolicy(iDx),gBRoundingUnit(iDx))
   End Select      

End Function

'===============================================================================
Function ValidDateCheck(pObjFromDt, pObjToDt)

	ValidDateCheck = False
	If Len(Trim(pObjToDt.Text)) Then
       If Len(Trim(pObjFromDt.Text)) Then
          If UniConvDateToYYYYMMDD(pObjFromDt.Text,pObjFromDt.UserDefinedFormat,"") > UniConvDateToYYYYMMDD(pObjToDt.Text,pObjToDt.UserDefinedFormat,"") Then
             Call DisplayMsgBox("970023","X", pObjToDt.Alt, pObjFromDt.Alt)
             Call SetFocusToDocument("M")
             pObjToDt.focus
             Set gActiveElement = document.activeElement
             Exit Function
          End If
       End If
	End If

	ValidDateCheck = True

End Function

