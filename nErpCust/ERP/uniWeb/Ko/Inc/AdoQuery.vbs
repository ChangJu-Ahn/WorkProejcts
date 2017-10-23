
'=======================================================================================
Function CommonQueryRs2by2(SelectList, FromList, WhereList, iRetArr)

    On Error Resume Next

    CommonQueryRs2by2 = False
    
    If gRdsUse = "T" Then
       CommonQueryRs2by2 = RDSCommonQueryRs2by2(SelectList, FromList, WhereList, iRetArr)
    Else
       CommonQueryRs2by2 = HTTPCommonQueryRs2by2(SelectList, FromList, WhereList, iRetArr)
    End If
    

End Function


'=======================================================================================
Function HTTPCommonQueryRs2by2(SelectList, FromList, WhereList, iRetArr)

    Dim ii, jj
    Dim iOutData
    Dim arrRow, arrCol
    
    On Error Resume Next

    HTTPCommonQueryRs2by2 = False
    
    iRetArr = ""

    If HTTPQuery(SelectList, FromList, WhereList, iOutData) = False Then
       Exit Function
    End If
    
    If IsEmpty(iOutData) Then
       Exit Function
    End If
    
    If Trim(iOutData) = "" Then
       Exit Function
    End If
    
    arrRow = Split(iOutData, Chr(12))

    For ii = 0 To UBound(arrRow) - 1
        arrCol = Split(arrRow(ii), Chr(11))
        
        For jj = 0 To UBound(arrCol) '- 1
            iRetArr = iRetArr & Chr(11) & arrCol(jj)
        Next
        iRetArr = iRetArr & Chr(11) & Chr(12)
    Next

    
    HTTPCommonQueryRs2by2 = True

End Function


'=======================================================================================
Function RDSCommonQueryRs2by2(SelectList, FromList, WhereList, iRetArr)

    Dim rs0, i, j
    
    On Error Resume Next

    RDSCommonQueryRs2by2 = False
    
    iRetArr = ""

    If RDSQuery(SelectList, FromList, WhereList, rs0) = False Then
       Exit Function
    End If

    If rs0 Is Nothing Then
       Exit Function
    End If

    If (IsNull(rs0)) Or (rs0 Is Nothing) Or (rs0.EOF And rs0.BOF) Then
       rs0.Close
       Set rs0 = Nothing
       Exit Function
    End If

    i = 0

    While Not rs0.EOF
          For j = 0 To rs0.Fields.Count - 1
              iRetArr = iRetArr & Chr(11) & rs0(j)
          Next
          i = i + 1
          iRetArr = iRetArr & Chr(11) & Chr(12)
          rs0.MoveNext
    Wend
    
    rs0.Close
    Set rs0 = Nothing
    
    RDSCommonQueryRs2by2 = True
End Function


'=======================================================================================
Function CommonQueryRs(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    On Error Resume Next
    
    CommonQueryRs = False
    
    lgF0 = ""
    lgF1 = ""
    lgF2 = ""
    lgF3 = ""
    lgF4 = ""
    lgF5 = ""
    lgF6 = ""

    If gRdsUse = "T" Then
       CommonQueryRs = RDSQueryMain(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    Else
       CommonQueryRs = HTTPQueryMain(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    End If
    
End Function

'=======================================================================================
Function HTTPQueryMain(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    Dim iOutData
    Dim arrRow, arrCol
    Dim ii
    Dim iiMax, jjMax
    
    Dim Tmp(6)

    On Error Resume Next
    
    HTTPQueryMain = False
    
    If HTTPQuery(SelectList, FromList, WhereList, iOutData) = False Then
       Exit Function
    End If
    
    If IsEmpty(iOutData) Then
       Exit Function
    End If
    
    If Trim(iOutData) = "" Then
       Exit Function
    End If
    
    arrRow = Split(iOutData, Chr(12))
    For ii = 0 To UBound(arrRow) - 1
        arrCol = Split(arrRow(ii), Chr(11))
        lgF0 = lgF0 & arrCol(0) & Chr(11)
        If UBound(arrCol) > 0 Then
           lgF1 = lgF1 & arrCol(1) & Chr(11)
           If UBound(arrCol) > 1 Then
              lgF2 = lgF2 & arrCol(2) & Chr(11)
              If UBound(arrCol) > 2 Then
                 lgF3 = lgF3 & arrCol(3) & Chr(11)
                 If UBound(arrCol) > 3 Then
                    lgF4 = lgF4 & arrCol(4) & Chr(11)
                    If UBound(arrCol) > 4 Then
                       lgF5 = lgF5 & arrCol(5) & Chr(11)
                       If UBound(arrCol) > 5 Then
                          lgF6 = lgF6 & arrCol(6) & Chr(11)
                       End If
                    End If
                 End If
              End If
           End If
       End If
    Next
    
    HTTPQueryMain = True

End Function

'=======================================================================================
Function RDSQueryMain(SelectList, FromList, WhereList, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

    Dim rs0
    On Error Resume Next
    
    RDSQueryMain = False

    If RDSQuery(SelectList, FromList, WhereList, rs0) = False Then
       Exit Function
    End If
    
    If (IsNull(rs0)) Or (rs0 Is Nothing) Or (rs0.EOF And rs0.BOF) Then
       rs0.Close
       Set rs0 = Nothing
       Exit Function
    End If
    
    While Not rs0.EOF
          If rs0.Fields.Count > 0 Then
             lgF0 = lgF0 & rs0(0) & Chr(11)
             If rs0.Fields.Count > 1 Then
                lgF1 = lgF1 & rs0(1) & Chr(11)
                If rs0.Fields.Count > 2 Then
                   lgF2 = lgF2 & rs0(2) & Chr(11)
                   If rs0.Fields.Count > 3 Then
                      lgF3 = lgF3 & rs0(3) & Chr(11)
                      If rs0.Fields.Count > 4 Then
                         lgF4 = lgF4 & rs0(4) & Chr(11)
                         If rs0.Fields.Count > 5 Then
                            lgF5 = lgF5 & rs0(5) & Chr(11)
                            If rs0.Fields.Count > 6 Then
                               lgF6 = lgF6 & rs0(6) & Chr(11)
                            End If  ' 6
                         End If  ' 5
                      End If  ' 4
                   End If  ' 3
                End If  ' 2
             End If  ' 1
          End If  ' 0
          rs0.MoveNext
    Wend

    rs0.Close
    Set rs0 = Nothing
    
    RDSQueryMain = True

End Function

'=======================================================================================
Function HTTPQuery(ByVal SelectList, ByVal FromList, ByVal WhereList, prData)
    Dim iStrSQL
    Dim iXmlHttp
    Dim iRetByte 
    
    On Error Resume Next
    Err.Clear
    
    HTTPQuery = False

    iStrSQL = "Select " & SelectList
    
    If Trim(FromList) > "" Then
       iStrSQL = iStrSQL & " From  " & FromList
       If Trim(WhereList) > "" Then
          iStrSQL = iStrSQL & " Where  " & WhereList
       End If
       
    End If

    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP")
    
    iXmlHttp.open "POST", GetComaspFolderPath & "RequestCommonQry.asp", False
    iXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    iStrSQL = Escape(iStrSQL)
    iStrSQL = Replace(iStrSQL, "+", "%2B")
    iStrSQL = Replace(iStrSQL, "/", "%2F")

    iXmlHttp.send "LangCD=" & gLang & "&ADODBConnString=" & Escape(gADODBConnString) & "&StrSQL=" & iStrSQL


    If gCharSet = "D" Then
       prData   = ConnectorControl.CStrConv(iXmlHttp.responseBody)
    Else
       prData   =                  iXmlHttp.responseText
    End If   
  
    Set iXmlHttp = Nothing
    If prData <> "" Then
        HTTPQuery = True
    End If
End Function

'=======================================================================================
Function RDSQuery(SelectList, FromList, WhereList, rs0)
    Dim ADF                                                                    '¢Ð : declaration Variable indicating ActiveX Data Factory
    Dim lgStrSQL
    Dim strRetMsg                                                              '¢Ð : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag                                   '¢Ð : declaration DBAgent Parameter

    On Error Resume Next

    Err.Clear
    
    ReDim UNISqlId(0)
    ReDim UNIValue(0, 0)
    
    RDSQuery = False
    
    lgStrSQL = "Select " & SelectList
    
    If Trim(FromList) > "" Then
       lgStrSQL = lgStrSQL & " From  " & FromList
       
       If Trim(WhereList) > "" Then
          lgStrSQL = lgStrSQL & " Where  " & WhereList
       End If
       
    End If

    UNISqlId(0) = "commonqry"
    UNIValue(0, 0) = lgStrSQL
    UNILock = DISCONNREAD: UNIFlag = "1"
    
    If Trim(gDsnNo) = "" Then
       Exit Function
    End If

    If Trim(gServerIP) = "" Then
       Exit Function
    End If

    Set ADF = ADS.CreateObject("prjPublic.cCtlTake", gServerIP)
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    If Err.Number <> 0 Then
       Set ADF = Nothing
       Exit Function
    End If

    RDSQuery = True

    Set ADF = Nothing

End Function
	  
Function FncGetEmpInf(pEmpNo,pRetStatus,pEmpName,pDeptNm,pRollPstn,pPayGrd1,pPayGrd2,pEntrDt)
    Dim ADF                                                                    '¢Ð : declaration Variable indicating ActiveX Data Factory
    Dim lgstrRetMsg                                                            '¢Ð : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '¢Ð : declaration DBAgent Parameter 
	
	Err.Clear

	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	FncGetEmpInf =  True

	pRetStatus = ""
	pEmpName   = ""
	pDeptNm    = ""
	pRollPstn  = ""
	pPayGrd1   = ""
	pPayGrd2   = ""
	pEntrDt    = ""	

    lgStrSQL =             "Select EMP_NO,NAME,DEPT_NM,ROLL_PSTN,PAY_GRD1,PAY_GRD2,ENTR_DT "
    lgStrSQL = lgStrSQL  & " From HAA010T "
    lgStrSQL = lgStrSQL  & " WHERE EMP_NO = '" & Replace(pEmpNo,"'","''") & "'"
    
	UNISqlId(0)    = "commonqry"
	UNIValue(0, 0) = lgStrSQL

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	If Trim(gDsnNo) ="" Then
       FncGetEmpInf = False
       Exit Function
	End If
	
	Set ADF = ADS.CreateObject("prjPublic.cCtlTake",gServerIP )
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If IsNull(rs0) Then
       FncGetEmpInf = False
       Exit Function
	End If
	
	If IsEmpty(rs0) Then
       FncGetEmpInf = False
       Exit Function
	End If

	If rs0.EOF And rs0.BOF Then
       FncGetEmpInf = False
 	   rs0.Close
       Set rs0 = Nothing
       Set ADF = Nothing			
       Exit Function
	End If
	
	pEmpName   = rs0(1)
	pDeptNm    = rs0(2)
	pRollPstn  = FuncCodeName(1,"H0002",rs0(3))
	pPayGrd1   = FuncCodeName(1,"H0001",rs0(4))
	pPayGrd2   = rs0(5)
	pEntrDt    = rs0(6)	
	
    rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing			

End Function
	  	  
	  	  
Function FuncCodeName(intSW, MajorCd, MinorCd)
    Dim iSelectList
    Dim iFromList
    Dim iWhereList
    
    
    
    Select Case intSW
        Case 1                                                  ' B_MAJOR
              iSelectList = " MINOR_NM "
              iFromList   = " B_MINOR  "
              iWhereList  = " MAJOR_CD = '" & MajorCd & "' AND MINOR_CD = '" & MinorCd & "'" 
              
        Case 2                                                  ' B_ACCT_DEPT  : dept
              iSelectList = " DEPT_NM "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & MinorCd & "')"
              Else
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
              End If   
        Case 3                                                  ' B_COUNTRY : ±¹Àû
              iSelectList = " COUNTRY_NM "
              iFromList   = " B_COUNTRY  "
              iWhereList  = " COUNTRY_CD = '" & MinorCd & "'"     

        Case 4                                                  ' B_COMPANY : company
              iSelectList = " CO_NM "
              iFromList   = " B_COMPANY  "
              iWhereList  = " CO_CD = '" & MinorCd & "'"  
        Case 5                                                  ' interanl dept
              iSelectList = " INTERNAL_CD "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & MinorCd & "')"   
              Else
                 iWhereList  = " DEPT_CD    = '" & MajorCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())" 
              End If 
	End Select

    If 	CommonQueryRs(iSelectList,iFromList,iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncCodeName = MinorCd
    Else
        lgF0 = Split(lgF0,Chr(11))
        FuncCodeName = lgF0(0)
    End If

End Function

Function FuncDeptName(DeptCd, OrgChangeDt, lgIntCd, DeptNm, IntCd)
    Dim iWhereList
    Dim strIntCd

    DeptNm = ""
    IntCd = ""

    If  OrgChangeDt > "" Then
	iWhereList = " DEPT_CD = '" & DeptCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & OrgChangeDt & "')"
    Else
        iWhereList = " DEPT_CD = '" & DeptCd & "' AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
    End If   

    If 	CommonQueryRs(" DEPT_NM,INTERNAL_CD "," B_ACCT_DEPT ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncDeptName = -1	' not exists dept table
	exit function
    end if

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))

    strIntCd = Trim(Replace(lgIntCd, "%", ""))

    if (strIntCd = "") OR (LEFT(lgF1(0), Len(strIntCd)) <> strIntCd) then
        FuncDeptName= -2	' no authority
    else
        DeptNm = Trim(lgF0(0))
        IntCd = Trim(lgF1(0))
	FuncDeptName= 0
    End If

End Function

Function FuncGetAuth(PgmId, UsrId, plgIntCd)

    If 	CommonQueryRs(" INTERNAL_CD,AUTH_YN "," HZA010T ", " MNU_ID = '" & UCase(PgmId) & "' AND USR_ID = '" & UCase(UsrId) & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	    plgIntCd = "1"		' find all authority if not exists in authority table
        FuncAuth = 0		' not exists in authority
	exit function
    End If

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))	' AUTH_YN

    if	Trim(lgF1(0)) = "N" then	'did not check authority
    	plgIntCd = "1"		'	find all authority if not exists in authority table
    else
	plgIntCd = Trim(lgF0(0))
    end if

    FuncAuth= 0

End Function

Function FuncGetEmpInf2(pEmpNo,plgIntCd,pEmpName,pDeptNm,pRollPstn,pPayGrd1,pPayGrd2,pEntrDt,pIntCd)
    Dim ADF                                                                    '¢Ð : declaration Variable indicating ActiveX Data Factory
    Dim lgstrRetMsg                                                            '¢Ð : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '¢Ð : declaration DBAgent Parameter 
    Dim strlgIntCd

    Err.Clear

    Redim UNISqlId(0)
    Redim UNIValue(0, 0)
	
    FuncGetEmpInf2 =  0

    pEmpName   = ""
    pDeptNm    = ""
    pRollPstn  = ""
    pPayGrd1   = ""
    pPayGrd2   = ""
    pEntrDt    = ""	
    pIntCd     = ""	

    lgStrSQL =             "Select EMP_NO,NAME,DEPT_NM,ROLL_PSTN,PAY_GRD1,PAY_GRD2,ENTR_DT,INTERNAL_CD "
    lgStrSQL = lgStrSQL  & " From HAA010T "
    lgStrSQL = lgStrSQL  & " WHERE EMP_NO = '" & Replace(pEmpNo,"'","''") & "'"
    
    UNISqlId(0)    = "commonqry"
    UNIValue(0, 0) = lgStrSQL

    UNILock = DISCONNREAD :	UNIFlag = "1"
	
    If  Trim(gDsnNo) ="" Then
    	FuncGetEmpInf2 = -1
       	Exit Function
    End If
	
    Set ADF = ADS.CreateObject("prjPublic.cCtlTake",gServerIP )
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    If  IsNull(rs0) Then
    	FuncGetEmpInf2 = -1
        Exit Function
    End If
	
    If  rs0 Is Nothing Then
	FuncGetEmpInf2 = -1
	Exit Function
    End If

    If  rs0.EOF And rs0.BOF Then
  	FuncGetEmpInf2 = -1
 	rs0.Close
        Set rs0 = Nothing
        Set ADF = Nothing			
        Exit Function
    End If

    strlgIntCd = Trim(Replace(plgIntCd, "%", ""))
    If strlgIntCd="" Then
        strlgIntCd="1"   ' set default value if a program did not need authority
    End If

    if (strlgIntCd = "") OR (LEFT(rs0(7), Len(strlgIntCd)) <> strlgIntCd) then
        FuncGetEmpInf2 = -2	' no authority
	exit function
    else	
	pEmpName   = Trim(rs0(1))
	pDeptNm    = Trim(rs0(2))
	pRollPstn  = Trim(FuncCodeName(1,"H0002",rs0(3)))
	pPayGrd1   = Trim(FuncCodeName(1,"H0001",rs0(4)))
	pPayGrd2   = Trim(rs0(5))
	pEntrDt    = Trim(rs0(6))
	pIntCd     = Trim(rs0(7))
        FuncGetEmpInf2 = 0	' normal
    end if

    rs0.Close
    Set rs0 = Nothing
    Set ADF = Nothing			

End Function


Function FuncGetEmpInf3(pEmpNo,plgIntCd,pEmpName,pDeptNm,pRollPstn,pPayGrd1,pPayGrd2,pEntrDt,pIntCd)
    Dim ADF                                                                    '¢Ð : declaration Variable indicating ActiveX Data Factory
    Dim lgstrRetMsg                                                            '¢Ð : declaration Variable indicating Record Set Return Message
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '¢Ð : declaration DBAgent Parameter 
    Dim strlgIntCd

    Err.Clear

    Redim UNISqlId(0)
    Redim UNIValue(0, 0)
	
    FuncGetEmpInf3 =  0

    pEmpName   = ""
    pDeptNm    = ""
    pRollPstn  = ""
    pPayGrd1   = ""
    pPayGrd2   = ""
    pEntrDt    = ""	
    pIntCd     = ""	

    lgStrSQL =             "Select EMP_NO,SUR_NAME,DEPT_NM,ROLL_PSTN,PAY_GRD1,FIRST_NAME,ENTR_DT,INTERNAL_CD "
    lgStrSQL = lgStrSQL  & " From H_HAA010T_SJP000 "
    lgStrSQL = lgStrSQL  & " WHERE EMP_NO = '" & pEmpNo & "'"
    
    UNISqlId(0)    = "commonqry"
    UNIValue(0, 0) = lgStrSQL

    UNILock = DISCONNREAD :	UNIFlag = "1"
	
    If  Trim(gDsnNo) ="" Then
    	FuncGetEmpInf3 = -1
       	Exit Function
    End If
	
    Set ADF = ADS.CreateObject("prjPublic.cCtlTake",gServerIP )
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    If  IsNull(rs0) Then
    	FuncGetEmpInf3 = -1
        Exit Function
    End If
	
    If  rs0 Is Nothing Then
	FuncGetEmpInf3 = -1
	Exit Function
    End If

    If  rs0.EOF And rs0.BOF Then
  	FuncGetEmpInf3 = -1
 	rs0.Close
        Set rs0 = Nothing
        Set ADF = Nothing			
        Exit Function
    End If

    strlgIntCd = Trim(Replace(plgIntCd, "%", ""))
    If strlgIntCd="" Then
        strlgIntCd="1"     ' set default value if a program did not need authority
    End If

    if (strlgIntCd = "") OR (LEFT(rs0(7), Len(strlgIntCd)) <> strlgIntCd) then
        FuncGetEmpInf3 = -2	' no authority
	exit function
    else	
	pEmpName   = Trim(rs0(1))
	pDeptNm    = Trim(rs0(2))
	pRollPstn  = Trim(FuncCodeName(1,"H0002",rs0(3)))
	pPayGrd1   = Trim(FuncCodeName(1,"H0001",rs0(4)))
	pPayGrd2   = Trim(rs0(5))
	pEntrDt    = Trim(rs0(6))
	pIntCd     = Trim(rs0(7))
        FuncGetEmpInf3 = 0	' normal
    end if

    rs0.Close
    Set rs0 = Nothing
    Set ADF = Nothing			

End Function




Function FuncGetTermDept(plgIntCd, pChngDt, rFrDept, rToDept)

    Dim iWhereList

    rFrDept = ""
    rToDept = ""

    If  pChngDt > "" Then
	iWhereList = " INTERNAL_CD LIKE '" & plgIntCd & "%' AND ORG_CHANGE_DT=(SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= '" & pChngDt & "')"
    Else
        iWhereList = " INTERNAL_CD LIKE '" & plgIntCd & "%' AND ORG_CHANGE_DT=(SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
    End If   

    If 	CommonQueryRs(" MIN(internal_cd), MAX(internal_cd) "," B_ACCT_DEPT ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncGetTermDept = -1	' not exists in dept table
	exit function
    end if

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))

    rFrDept = Trim(lgF0(0))
    rToDept = Trim(lgF1(0))

    FuncGetTermDept = 0

End Function

Function FuncLastMonthDay(pDate, rDate)

   Dim strDate1
   Dim strDate2

   strDate1 = Trim(Replace(pDate, gComDateType, ""))
   if strDate1 = "" then
      strDate2 = Year(Date) & gComDateType & Right("0" & Month(Date),2)
   else
      strDate2 = Mid(strDate1, 1, 4) & gComDateType
      strDate2 = strDate2 & Mid(strDate1, 5, 2)
      strDate2 = strDate2 & gComDateType & "01"
   end if

   rDate = DateAdd("D",-1, DateAdd("M",1,strDate2))

   FuncLastMonthDay = Day(rDate)

End Function