<%
'==== I_MKSCRIPT
'==== gADODBConnString
'==== gAPDateFormat
'==== gAPDateSeperator
'==== gClientDateFormat
'==== gClientDateSeperator
'==== gComDateType
'==== gDateFormat
'==== gDateFormatYYYYMM
'==== Function DisplayMsgBox

'==============================================================================
' SQL Server Date Format
'==============================================================================
Const gServerDateFormat = "YYYY-MM-DD"	' Server date format 
Const gServerBaseDate   = "1900-01-01"	' Server base date
Const gServerDateType   = "-"		    ' Seperator for server date

'==============================================================================
' Desc    : Date conversion 
' Rule    : Company -> DataBase(AP,YYYY-MM-DD)
'==============================================================================
Function UNIConvDate(ByVal pDate)
	Dim strYear, strMonth, strDay

  '  On Error Resume Next

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

    Call ExtractDateFrom(pDate,gDateFormat,gComDateType,strYear,strMonth,strDay)   ' From Company Date Type
    
    UNIConvDate = strYear & gServerDateType & strMonth & gServerDateType & strDay  ' To Server Date Type
    
End Function

'==============================================================================
' Desc    : Date conversion (YMD)
' Rule    : DataBase(AP) -> Company
'==============================================================================
Function UNIDateClientFormat(ByVal pDate)
'    On Error Resume Next		' 2002년 10월30일 영업모듈 조인국 요청에의해서 뺌 
    
	UNIDateClientFormat =  UNIDateClientFormatSub(pDate,"YMD")
End Function

'==============================================================================
' Desc    : Date conversion (YM)
' Rule    : DataBase -> Company
'==============================================================================
Function UNIMonthClientFormat(ByVal pDate)
    On Error Resume Next
    
    UNIMonthClientFormat = UNIDateClientFormatSub(pDate,"YM")
End Function

'==============================================================================
' Rule    : DataBase(AP) -> Company
'==============================================================================
Function UNIDateClientFormatSub(ByVal pDate,ByVal pOption)

	Dim strYear, strMonth, strDay
	Dim iTempDate
	
'    On Error Resume Next			' 2002년 10월30일 영업모듈 조인국 요청에의해서 뺌 
    
	UNIDateClientFormatSub = ""
	
	If IsNull(pDate) Or Trim(pDate) = "" Then
	   Exit Function 
	End If

    strYear  =              Year(pDate)
    strMonth = Right("0" & Month(pDate) ,2)
    strDay   = Right("0" &   Day(pDate) ,2)
    
    iTempDate = strYear & gServerDateType & strMonth & gServerDateType & strDay 
    
	If iTempDate <= gServerBaseDate Then
       Exit Function
	End If
	
    UNIDateClientFormatSub = MakeDateTo(pOption,gDateFormat,gComDateType,strYear,strMonth,strDay)

End Function

'==============================================================================
' Desc    : Date conversion (YMD)
' Rule    : DataBase(AP) -> Company
'==============================================================================
Function UNIConvDateDBToCompany(ByVal pDate, ByVal pDefault)
    On Error Resume Next

    UNIConvDateDBToCompany = ""
    
    If Trim(pDate) = "" Or IsNull(pDate) Then
       pDate = pDefault
       If Trim(pDate) = "" Or IsNull(pDate) Then
          Exit Function
       End If
    End If

   UNIConvDateDBToCompany = UNIDateClientFormatSub(pDate,"YMD")
End Function


'==============================================================================
' Desc    : Date conversion 
' Rule    : Company -> DataBase(AP)
'==============================================================================
Function UNIConvDateCompanyToDB(ByVal pDate,ByVal pDefault)

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
' Function Name : UniConvDateToYYYYMMDD
' Function Desc : Convert Date(?  -> YYYY-MM-DD)
'======================================================================================================
Function UniConvDateToYYYYMMDD(ByVal pDate , ByVal pDateFormat , ByVal pDateSeperator)
	Dim strYear, strMonth, strDay

    On Error Resume Next
    
    UniConvDateToYYYYMMDD = ""
    
    If IsNull(pDate) Or Trim(pDate) = "" Then
       Exit Function
    End If

    pDate = FillLeadingSpaceWithZero(pDate,pDateFormat)

    Call ExtractDateFromSuper(pDate,pDateFormat,strYear,strMonth,strDay)

    UniConvDateToYYYYMMDD = strYear & pDateSeperator & strMonth  & pDateSeperator & strDay

End Function

'======================================================================================================
' Function Name : UniConvYYYYMMDDToDate
' Function Desc : Convert Date( YYYY-MM-DD -> ? )
'======================================================================================================
Function UniConvYYYYMMDDToDate(ByVal pDateFormat ,ByVal strYear,ByVal strMonth, ByVal strDay)
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
' Function Name : UniConvDateToYYYYMM
' Function Desc : Convert Date(?  -> YYYY-MM)
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
' Function Name : UniConvDateAToB
' Function Desc : Convert Date(?  -> ?)
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
' Function Name : UNIDateAdd
' Function Desc : UNIDateAdd is same sa DateAdd
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
	
    mDate = FillLeadingSpaceWithZero(mDate,gAPDateFormat)
    
    UNIDateAdd = UniConvDateAToB(mDate,gAPDateFormat,pDateFormat)
    
End Function

'===============================================================================
' Function Name : CompareDateByFormat
' Function Desc : Compare Date
'===============================================================================
Function CompareDateByFormat(pFromDt, pToDt,pFromDtAlt,pToDtAlt,pMsgCD,pDateFormat,pDateSeperator,pBool)

    Dim strYear1
    Dim strMonth1
    Dim strDay1
    Dim strFullDay1
    
    Dim strYear2
    Dim strMonth2
    Dim strDay2
    Dim strFullDay2

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
                   Call DisplayMsgBox(pMsgCD,vbOKOnly, pToDtAlt   , pFromDtAlt,I_MKSCRIPT)
                Else
                   Call DisplayMsgBox(pMsgCD,vbOKOnly, pFromDtAlt , pToDtAlt  ,I_MKSCRIPT)
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
Function CompareDateByFormat2(pFromDt, pToDt,pFromDtAlt,pToDtAlt,pMsgCD,pDateFormat,pDateSeperator,pBool)

    Dim strYear1
    Dim strMonth1
    Dim strDay1
    Dim strFullDay1
    
    Dim strYear2
    Dim strMonth2
    Dim strDay2
    Dim strFullDay2

	CompareDateByFormat2 = False
    
    Call ExtractDateFrom(pFromDt,pDateFormat,pDateSeperator,strYear1,strMonth1,strDay1)
    Call ExtractDateFrom(pToDt  ,pDateFormat,pDateSeperator,strYear2,strMonth2,strDay2)
      
    strFullDay1 = strYear1 & strMonth1 '& strDay1
    strFullDay2 = strYear2 & strMonth2 '& strDay2

	If Len(Trim(strFullDay2)) Then
       If Len(Trim(strFullDay1)) Then
          If (strFullDay1 + 11 ) > strFullDay2 Then
			Else
             'If pBool = True Then
                'If pMsgCD = "970023" Then
                   'Call DisplayMsgBox(pMsgCD,vbOKOnly, pToDtAlt   , pFromDtAlt,I_MKSCRIPT)
                'Else
                   'Call DisplayMsgBox(pMsgCD,vbOKOnly, pFromDtAlt , pToDtAlt  ,I_MKSCRIPT)
                'End If   
             'End If   
            Exit Function
          End If
       End If
	End If

	CompareDateByFormat2 = True

End Function

'======================================================================================================
' Function Name : UNIGetLastDay
' Function Desc : get last day of this month
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
	
    mDate = FillLeadingSpaceWithZero(mDate,gAPDateFormat)
    
    If pDateFormat = gDateFormatYYYYMM Then
	   UNIGetLastDay = UniConvDateAToB(mDate,gAPDateFormat,gDateFormat)
	Else   
	   UNIGetLastDay = UniConvDateAToB(mDate,gAPDateFormat,pDateFormat)
	End If   

End Function

'======================================================================================================
' Function Name : UNIGetFirstDay
' Function Desc : get first day of this month
'======================================================================================================
Function UNIGetFirstDay(ByVal pDate, ByVal pDateFormat)
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
' Desc : Check If date format is server date type
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
    
    xDate = Replace(xDate,"JAN","XXX")
    xDate = Replace(xDate,"FEB","XXX")
    xDate = Replace(xDate,"MAR","XXX")
    xDate = Replace(xDate,"APR","XXX")
    xDate = Replace(xDate,"MAY","XXX")
    xDate = Replace(xDate,"JUN","XXX")
    xDate = Replace(xDate,"JUL","XXX")
    xDate = Replace(xDate,"AUG","XXX")
    xDate = Replace(xDate,"SEP","XXX")
    xDate = Replace(xDate,"OCT","XXX")
    xDate = Replace(xDate,"NOV","XXX")
    xDate = Replace(xDate,"DEC","XXX")
    
    xDateFormat = Replace(pDateFormat,"Y","X")
    xDateFormat = Replace(xDateFormat,"M","X")
    xDateFormat = Replace(xDateFormat,"D","X")
    
    xDateSeperator = Replace(xDateFormat,"X","")

    xDateSeperator = Mid(xDateSeperator,1,1)
    
    cDateFormatArr = Split(pDateFormat,xDateSeperator)
    
    If Instr(pDate,xDateSeperator) > 0 Then
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
' Desc : Extract year,month,day from date
'==============================================================================
Sub ExtractDateFromSuper(ByVal pDate,pDateFormat,strYear,strMonth,strDay)

    If ValidateData(pDate,"SEN") = False Then   '2002/09/28 jinsoo lee
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
' Desc : Extract year,month,day from date
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
Function MakeDateTo(ByVal pOpt,ByVal pDateFormat,ByVal pDateSeperator,ByVal strYear,ByVal strMonth,ByVal strDay)

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

'-----------------------------------------------------------------------------------------------------------------------
' This function return Date + Time
' If Date fomrat is like "2001-1-2", Then return reformated date ("2001-01-02")
'-----------------------------------------------------------------------------------------------------------------------
Function GetSvrDateTime()
    
	Dim adoRec
	Dim iSTRSQL
	
'   1234567890123456789
'   2002-02-19 20:19:04.043 

    On Error Resume Next
    
    If gDBKind = "ORACLE" Then  
       iSTRSQL = "SELECT to_char(sysdate,'YYYY-MM-DD HH24:MI:SS') FROM DUAL "                                                                      
    Else 
       iSTRSQL = "SELECT CONVERT(CHAR(24),GetDate(), 21)"                                                                      
    End If   

	Set adoRec = Server.CreateObject("ADODB.RecordSet")    
	                                    ' adOpenForwardOnly, adLockReadOnly, adCmdTable
    adoRec.Open iSTRSQL,gADODBConnString, 0                , 1             , 1

	If Err.Number = 0 Then
	   If Not IsNull(adoRec(0)) Then
          GetSvrDateTime = adoRec(0)
       End If
       adoRec.Close
       Set adoRec = Nothing
    End If   

    GetSvrDateTime = Trim(GetSvrDateTime)

    If GetSvrDateTime = "" Then
       GetSvrDateTime = gServerBaseDate & " 1:01:01"
       Exit Function
    Else
       GetSvrDateTime = Mid(GetSvrDateTime,1,19)    
    End If

End Function

'==============================================================================
'
'==============================================================================
Function GetSvrDate()
    On Error Resume Next

	GetSvrDate = Mid(GetSvrDateTime, 1, 10)
End Function

'==============================================================================
'
'==============================================================================
Function GetSvrDateYYYYMM()
    On Error Resume Next

	GetSvrDateYYYYMM = Mid(GetSvrDateTime, 1, 7)
End Function

'==============================================================================
'
'==============================================================================
Function GetSvrTime()
    On Error Resume Next

	GetSvrTime = Trim(Mid(GetSvrDateTime, 11))
End Function

'==============================================================================
' Desc : Fill Leading Space with Zero
'==============================================================================
Function FillLeadingSpaceWithZero(ByVal pDate,ByVal pDateFormat)
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
' Desc : convert yy to yyyy
'==============================================================================
Function ConvertYYToYYYY(ByVal pYY)

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


'==============================================================================
' Name : ValidateData
' Desc : Validate Data
' Born : 2002/08/22
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

%>