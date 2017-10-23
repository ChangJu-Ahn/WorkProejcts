<%

'-----------------------------------------------------------------------------------------------------------------------
' This function return Date + Time
' If Date fomrat is like "2001-1-2", Then return reformated date ("2001-01-02")
'-----------------------------------------------------------------------------------------------------------------------
Function GetSvrDateTime()
    
	Dim adoRec
	Dim iSTRSQL
	
'   1234567890123456789
'   2002-02-19 20:19:04.043 

    On Error Resume Next												            '☜: Server Side Process

    If GetGlobalData("gDBKind") = "ORACLE" Then  
       iSTRSQL = "SELECT to_char(sysdate,'YYYY-MM-DD HH24:MI:SS') FROM DUAL "                                                                      
    Else 
       iSTRSQL = "SELECT CONVERT(CHAR(24),GetDate(), 21)"                                                                      
    End If   

	Set adoRec = Server.CreateObject("ADODB.RecordSet")    
	                                    ' adOpenForwardOnly, adLockReadOnly, adCmdTable
    adoRec.Open iSTRSQL,GetGlobalData("gADODBConnString"), 0                , 1             , 1
	
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
Function GetSvrDate()
    On Error Resume Next

	GetSvrDate = Mid(GetSvrDateTime, 1, 10)

End Function

'==============================================================================
Function GetSvrDateYYYYMM()
    On Error Resume Next

	GetSvrDateYYYYMM = Mid(GetSvrDateTime, 1, 7)
End Function

'==============================================================================
Function GetSvrTime()
    On Error Resume Next

	GetSvrTime = Trim(Mid(GetSvrDateTime, 11))
End Function

'==============================================================================
Sub SubFillRemBodyTD5656(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD> </TR>"
   Next    
End Sub

'==============================================================================
Sub SubFillRemBodyTD656(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD656 NOWRAP>&nbsp;</TD></TR>"
   Next    
End Sub

'==============================================================================
Sub SubFillRemBodyTD656C(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD656 COLSPAN=3 NOWRAP>&nbsp;</TD></TR>"
   Next    
End Sub

'==============================================================================
Sub SubFillRemBodyTD56(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD></TR>"
   Next    
End Sub

Function CommonQueryRs(SelectList,FromList,WhereList)

    Dim lgStrSQL
    Dim pRs
    
    On Error Resume Next

    CommonQueryRs = ""

    lgStrSQL = "Select " & SelectList & " From   " & FromList

    If Trim(WhereList) > "" Then
       lgStrSQL = lgStrSQL  & " Where  " & WhereList
    End If
    
    Set pRs = Server.CreateObject("ADODB.Recordset")
    
    pRs.Open lgStrSQL,GetGlobalData("gADODBConnString"),0,1
    
    If Err.number <> 0  Then
    Else
       CommonQueryRs = pRs.GetString(,,Chr(11),Chr(12))
    End If
    
	If pRs.State = 1 then		' 연결중이면 연결 해재 
	   pRs.Close
       Set pRs = Nothing
	End If

End Function

Function GetGlobalData(pData)   '2003-08-07 leejinsoo
    Dim FileNm
    Dim xmlDoc
    
    On Error Resume Next

    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = False 
	    
    xmlDoc.LoadXML (GetSessionStream)

    GetGlobalData = xmlDoc.selectSingleNode("/uniERP/LoadBasisGlobalInf/" & pData ).text

    Set xmlDoc    = Nothing

End Function

Function GetSessionStream()

    Dim xmlDoc
    Dim xSessionDll
    
    On Error Resume Next

    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	Set xSessionDll = Server.CreateObject("xSession.A00001")
	xmlDoc.async = False 
	GetSessionStream = xSessionDll.DMakeDic(Request.Cookies("unierp")("SessionKey"))	
	Set xSessionDll = Nothing
    Set xmlDoc      = Nothing

End Function


Sub MakeComboData(SelectList,FromList,WhereList)

     Dim pRec

     Set pRec  = Server.CreateObject("ADODB.RecordSet")
     
     pRec.Open "SELECT " & SelectList & " from " & FromList & " WHERE " & WhereList ,GetGlobalData("gADODBConnString")
     
     If Err.number = 0 Then

        If pRec.EOF Or pRec.BOF Then
        Else
        '  Response.Write "<OPTION  SELECTED value= " & " >"
          Do While  Not (pRec.EOF Or pRec.BOF)
             Response.Write "<OPTION  value=" & pRec(0) & " >" & pRec(1)
             pRec.MoveNext
          Loop
        End If  
     End If   

     pRec.Close
     Set pRec  = Nothing


End Sub



Call MakeLocalAuth()

Sub MakeLocalAuth()

     Dim pRec

     Set pRec  = Server.CreateObject("ADODB.RecordSet")
     
     pRec.Open " select BA, PL, SG, SO, PG, PO from VIEW_Z_USR_ORG_MAST_KO441 where USR_ID = '" & GetGlobalData("gUsrId") & "'" ,GetGlobalData("gADODBConnString")
     
     If Err.number = 0 Then

        If pRec.EOF Or pRec.BOF Then
        Else

           Response.Write "<SCRIPT LANGUAGE=VBScript>" & vbCrLf
           Response.Write "Dim lgDefaultBA    " & vbCrLf	' -- 조직코드
           Response.Write "Dim lgDefaultPL      " & vbCrLf	' -- 부서코드
           Response.Write "Dim lgDefaultSG    " & vbCrLf	' -- Agent번호
           Response.Write "Dim lgDefaultSO    " & vbCrLf	' -- Agent명
           Response.Write "Dim lgDefaultPG    " & vbCrLf	' -- Agent명
           Response.Write "Dim lgDefaultPO    " & vbCrLf	' -- Agent명
           

           Response.Write " lgDefaultBA   = """ & pRec(0) & """" & vbCrLf
           Response.Write " lgDefaultPL     = """ & pRec(1) & """" & vbCrLf
           Response.Write " lgDefaultSG   = """ & pRec(2) & """" & vbCrLf
           Response.Write " lgDefaultSO   = """ & pRec(3) & """" & vbCrLf
           Response.Write " lgDefaultPG   = """ & pRec(4) & """" & vbCrLf
           Response.Write " lgDefaultPO   = """ & pRec(5) & """" & vbCrLf
          
           Response.Write "</SCRIPT>" & vbCrLf
        End If  
     End If   

     pRec.Close

     Set pRec  = Nothing


End Sub


%>