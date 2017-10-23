<%

Function CompanyAndDdList()
    
    Dim PCUniCommA
    Dim strCompanyList
    Dim iLoop
    Dim strDBList
    Dim iTotalList
    
    Set PCUniCommA = Server.CreateObject("PCUniComm.CA01")
    
    iTotalList = ""

    strCompanyList = Split(PCUniCommA.CompanyList, "||")
    
    If UBound(strCompanyList) <> -1 Then
       For iLoop = 0 To UBound(strCompanyList) - 1
           strDBList = Replace(PCUniCommA.DBList(Trim(strCompanyList(iLoop))), "||", ":")
           iTotalList = iTotalList & strCompanyList(iLoop) & ";" & strDBList & Chr(12)
       Next
    End If

    CompanyAndDdList = iTotalList
        
    Set PCUniCommA = Nothing

End Function


Function CompanyAndDdList2()
    
    Dim PCUniCommB
    Dim iDDB
    Dim iDDBNM
    Dim strCompanyList
    Dim iLoop
    Dim iTotalList
    
    Set PCUniCommB = Server.CreateObject("PCUniComm.CA01")
    
    iTotalList = ""

    strCompanyList = Split(PCUniCommB.CompanyList, "||")
    
    If UBound(strCompanyList) <> -1 Then
       For iLoop = 0 To UBound(strCompanyList) - 1
           iDDB   = Trim(PCUniCommB.ReadDefaultDB(Trim(strCompanyList(iLoop))))
           iDDBNM = Trim(PCUniCommB.ReadDefaultCompanyName(Trim(strCompanyList(iLoop))))
           iDDBNM = Replace(iDDBNM, Chr(0), "")
           
           If iDDBNM = "" Then
              iDDBNM = iDDB
           End If
           iTotalList = iTotalList & Trim(strCompanyList(iLoop)) & "<>" & iDDB  & "<>" & Trim(iDDBNM) & "::"
       Next
        
    End If

    CompanyAndDdList2 = iTotalList
        
    Set PCUniCommB = Nothing

End Function


%>