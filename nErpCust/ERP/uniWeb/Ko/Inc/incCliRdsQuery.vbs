
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

    If gCharSet = "D" Then 'U : unicode, D:DBCS
       prData   = ConnectorControl.CStrConv(iXmlHttp.responseBody)
    Else
       prData   = iXmlHttp.responseText
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