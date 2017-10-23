<% Option Explicit%>
<!-- #Include file="../inc/CommResponse.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<%
   Dim gOPT
   Dim gADODBConnString
   Dim gUSERID
   Dim adoRs
   Dim AdoCnn   
   Dim iSQL
   Dim iSTR
   Dim iSTRDefaultXML
   
   DIm gTimerOperationYN
   DIm gAllowTimerOperationYN
   DIm gTimeOutValue
   Dim gMinimumTimeOutValue
   Dim gMaximumTimeOutValue
   
   Dim gDefaultFontName
   Dim gDefaultFontSize
   Dim gIEReOpenYN
   Dim gKeepSessionAliveTime
   
   Dim gMWindowW
   Dim gMWindowH
   
   Dim gLang
   Dim gMsgText
   
   On Error Resume Next
   
   gOPT             = Request("iOPT")
   gADODBConnString = Request("iADODBConnString")
   gUSERID          = Request("iUSERID")

   Response.ContentType = "text/html"

   gLang = Request("LangCD")

   Call GetTimerInfo(gTimerOperationYN,gAllowTimerOperationYN,gTimeOutValue,gMinimumTimeOutValue,gMaximumTimeOutValue)
   Call GetuniConnectorInfo(gDefaultFontName,gDefaultFontSize,gIEReOpenYN,gKeepSessionAliveTime)
   Call GetuniConnectorInfo2(gMWindowW,gMWindowH)
   
   Select Case gOPT
        Case "Q"
            
            If ExistsTable("Z_CONNECTOR_CONFIG") Then   
            
               Set adoRs = Server.CreateObject("ADODB.Recordset")

               iSQL = "select isnull(FONT_NAME,'" & gDefaultFontName & "') FONT_NAME ,isnull(FONT_SIZE," & gDefaultFontSize & ") FONT_SIZE ,isnull(TIMEOUT,0) TIMEOUT, upper(isnull(REOPENIE,'Y')) REOPENIE from Z_CONNECTOR_CONFIG where USR_ID = '" & gUSERID & "'"

               Call WriteToLog(iSQL)

               adoRs.Open iSQL ,gADODBConnString

               If Err.number = 0 Then
                  If adoRs.EOF And adoRs.BOF Then                     
                              
                  Else
                     gTimeOutValue     = adoRs("TIMEOUT") 

                     gDefaultFontName  = adoRs("FONT_NAME") 
                     gDefaultFontSize  = adoRs("FONT_SIZE") 
                     
                     gIEReOpenYN       = adoRs("REOPENIE") 
                  End If
               End If
            End If  

            If adoRS.State = 1 Then
                adoRS.Close
            End If
            Set adoRs = Nothing       

            iSTRDefaultXML =                  "<uniERP>"  
            iSTRDefaultXML = iSTRDefaultXML &   "<TIMEOPERYN>"      & gTimerOperationYN      & "</TIMEOPERYN>"
            iSTRDefaultXML = iSTRDefaultXML &   "<ALLOWTIMEOPERYN>" & gAllowTimerOperationYN & "</ALLOWTIMEOPERYN>"
            iSTRDefaultXML = iSTRDefaultXML &   "<TIMEOUT>"         & gTimeOutValue          & "</TIMEOUT>"
            iSTRDefaultXML = iSTRDefaultXML &   "<TIMEOUTMIN>"      & gMinimumTimeOutValue   & "</TIMEOUTMIN>"
            iSTRDefaultXML = iSTRDefaultXML &   "<TIMEOUTMAX>"      & gMaximumTimeOutValue   & "</TIMEOUTMAX>"
            
            iSTRDefaultXML = iSTRDefaultXML &   "<FONTNAME>"        & gDefaultFontName       & "</FONTNAME>"
            iSTRDefaultXML = iSTRDefaultXML &   "<FONTSIZE>"        & gDefaultFontSize       & "</FONTSIZE>"
            
            iSTRDefaultXML = iSTRDefaultXML &   "<MWINDOWW>"        & gMWindowW              & "</MWINDOWW>"
            iSTRDefaultXML = iSTRDefaultXML &   "<MWINDOWH>"        & gMWindowH              & "</MWINDOWH>"

            iSTRDefaultXML = iSTRDefaultXML &   "<REOPENIE>"        & gIEReOpenYN            & "</REOPENIE>"
            iSTRDefaultXML = iSTRDefaultXML &   "<KSAT>"            & gKeepSessionAliveTime  & "</KSAT>"
            iSTRDefaultXML = iSTRDefaultXML & "</uniERP>"

            Response.Write iSTRDefaultXML


        Case "UA"

            Set adoRs = Server.CreateObject("ADODB.Recordset")
            Set AdoCnn = Server.CreateObject("ADODB.Connection")
            AdoCnn.Open gADODBConnString                  
            
            iSQL = "select TIMEOUT from Z_CONNECTOR_CONFIG where USR_ID = '" & gUSERID & "'"
            
            Call WriteToLog(iSQL)
            adoRs.Open iSQL , gADODBConnString

            If Err.number = 0 Then
               If adoRs.EOF And adoRs.BOF Then
                  iSQL = iSQL & " Insert into Z_CONNECTOR_CONFIG(USR_ID,FONT_NAME,FONT_SIZE,TIMEOUT,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values "
                  iSQL = iSQL & " ('" & gUSERID & "','" & gDefaultFontName & "'," & gDefaultFontSize & "," & Request("iTimeOut") & ",'" & gUSERID & "',getdate(),'" & gUSERID & "',getdate())"
               Else
                  iSQL = "update  Z_CONNECTOR_CONFIG set "
                  iSQL = iSQL & " TIMEOUT      =  " & Request("iTimeOut")  & "  , "
                  iSQL = iSQL & " UPDT_USER_ID = '" & gUSERID & "' , "
                  iSQL = iSQL & " UPDT_DT      = getdate() "
                  iSQL = iSQL & " where USR_ID = '" & gUSERID & "'"                 
               End If
            End If
  
            Call WriteToLog(iSQL)
            AdoCnn.Execute iSQL
            
            adoRs.Close
            Set adoRs = Nothing       

            AdoCnn.Close
            Set AdoCnn = Nothing       


        Case "UB"

            Set adoRs = Server.CreateObject("ADODB.Recordset")
            Set AdoCnn = Server.CreateObject("ADODB.Connection")
            AdoCnn.Open gADODBConnString                  
            
            adoRs.Open "select TIMEOUT from Z_CONNECTOR_CONFIG where USR_ID = '" & gUSERID & "'",gADODBConnString

            If Err.number = 0 Then
               If adoRs.EOF And adoRs.BOF Then
                  iSQL = iSQL & " Insert into Z_CONNECTOR_CONFIG(USR_ID,FONT_NAME,FONT_SIZE,REOPENIE,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values "
                  iSQL = iSQL & " ('" & gUSERID & "','" & gDefaultFontName & "'," & gDefaultFontSize & ",'" & Request("iREOPENIE") & "','" & gUSERID & "',getdate(),'" & gUSERID & "',getdate())"
               Else
                  iSQL = "update  Z_CONNECTOR_CONFIG set "
                  iSQL = iSQL & " REOPENIE     = '" & Request("iREOPENIE") & "' , " 
                  iSQL = iSQL & " UPDT_USER_ID = '" & gUSERID & "' , "
                  iSQL = iSQL & " UPDT_DT      = getdate() "
                  iSQL = iSQL & " where USR_ID = '" & gUSERID & "'"                 
               End If
            End If
            
            Call WriteToLog(iSQL)
            
            AdoCnn.Execute iSQL
            
            adoRs.Close
            Set adoRs = Nothing       

            AdoCnn.Close
            Set AdoCnn = Nothing       



        Case "H"
        

            Set adoRs = Server.CreateObject("ADODB.Recordset")
   
            iSTRDefaultXML = ""
            
            adoRs.Open "SELECT isnull(A.INTERFACE_ID,'') INTERFACE_ID, isnull(A.PWD,'') PWD, isnull(B.UID,'') UID, isnull(B.PASSWORD,'') PASSWORD FROM Z_USR_MAST_REC A LEFT OUTER JOIN E11002T  B ON A.INTERFACE_ID = B.UID WHERE A.USR_ID = '" & gUSERID & "'",gADODBConnString

            If Err.number = 0 Then
               If adoRs.EOF And adoRs.BOF Then
               Else
                  iSTRDefaultXML =                  "<uniERP>"  
                  iSTRDefaultXML = iSTRDefaultXML &   "<INTERFACE_ID>" & adoRs("INTERFACE_ID")   & "</INTERFACE_ID>"
                  iSTRDefaultXML = iSTRDefaultXML &   "<PWD>"          & trim(adoRs("PWD"))      & "</PWD>"
                  iSTRDefaultXML = iSTRDefaultXML &   "<UID>"          & adoRs("UID")            & "</UID>"
                  iSTRDefaultXML = iSTRDefaultXML &   "<PASSWORD>"     & trim(adoRs("PASSWORD")) & "</PASSWORD>"
                  iSTRDefaultXML = iSTRDefaultXML & "</uniERP>"
               End If
            End If
  
            Response.Write iSTRDefaultXML
            adoRs.Close
            Set adoRs = Nothing       

        Case "B"

               Set adoRs = Server.CreateObject("ADODB.Recordset")
               adoRs.Open "Select MSG_TEXT FROM B_MESSAGE WHERE MSG_CD  = '" & Request("iMSGCD") & "' AND   LANG_CD = '" & gLang & "'" ,gADODBConnString

               If Err.number = 0 Then
                  If adoRs.EOF And adoRs.BOF Then               
                  Else
                     gMsgText = adoRs("MSG_TEXT") 
                  End If
               End If   

               If adoRS.State = 1 Then
                  adoRS.Close
               End If
               Set adoRs = Nothing       
                  
               Response.Write gMsgText
            
   End Select               


Sub GetTimerInfo(pTimerOperationYN, pAllowTimerOperationYN, pTimeOutValue, pMiniMumTimeOutValue, pMaxiMumTimeOutValue)
   
    Dim pRecX
    Dim iSQL

    pTimerOperationYN = "Y"
    pAllowTimerOperationYN = "Y"
    pTimeOutValue = 30
    pMiniMumTimeOutValue = 1
    pMaxiMumTimeOutValue = 60
            
    Set pRecX = CreateObject("ADODB.Recordset")
    
    iSQL = "select upper(isnull(REFERENCE,'" & pTimerOperationYN & "')) from b_configuration where MAJOR_CD = 'Z0050' and MINOR_CD = '1' and SEQ_NO = 1 "
    Call WriteToLog(iSQL)
    pRecX.Open iSQL, gADODBConnString
            
    If Err.Number = 0 Then
       If pRecX.EOF And pRecX.BOF Then
       Else
          pTimerOperationYN = pRecX(0)
       End If
    Else
    End If
            
    If pRecX.State = 1 Then
       pRecX.Close
    End If
            
    pRecX.Open "select isnull(REFERENCE,'" & pAllowTimerOperationYN & "') from b_configuration where MAJOR_CD = 'Z0050' and MINOR_CD = '2' and SEQ_NO = 1 ", gADODBConnString
     
    If Err.Number = 0 Then
       If pRecX.EOF And pRecX.BOF Then
       Else
          pAllowTimerOperationYN = pRecX(0)
       End If
    Else
    End If
            
    If pRecX.State = 1 Then
       pRecX.Close
    End If
            
    pRecX.Open "select isnull(REFERENCE,'" & pTimeOutValue & "') from b_configuration where MAJOR_CD = 'Z0050' and MINOR_CD = '3' and SEQ_NO = 1 ", gADODBConnString
     
    If Err.Number = 0 Then
       If pRecX.EOF And pRecX.BOF Then
       Else
          pTimeOutValue = pRecX(0)
       End If
    Else
    End If
            
    If pRecX.State = 1 Then
       pRecX.Close
    End If
            
    pRecX.Open "select isnull(REFERENCE,'" & pMiniMumTimeOutValue & "') from b_configuration where MAJOR_CD = 'Z0050' and MINOR_CD = '4' and SEQ_NO = 1 ", gADODBConnString
     
    If Err.Number = 0 Then
       If pRecX.EOF And pRecX.BOF Then
       Else
          pMiniMumTimeOutValue = pRecX(0)
       End If
    Else
    End If
            
    If pRecX.State = 1 Then
       pRecX.Close
    End If
            
    pRecX.Open "select isnull(REFERENCE,'" & pMaxiMumTimeOutValue & "') from b_configuration where MAJOR_CD = 'Z0050' and MINOR_CD = '5' and SEQ_NO = 1 ", gADODBConnString
     
    If Err.Number = 0 Then
       If pRecX.EOF And pRecX.BOF Then
       Else
          pMaxiMumTimeOutValue = pRecX(0)
       End If
    Else
    End If
            
    If pRecX.State = 1 Then
       pRecX.Close
    End If
            
    Set pRecX = Nothing
            
End Sub





Sub GetuniConnectorInfo(pFontName, pFontSize, pIEReOpenYN, pKeepSessionAliveTime)
   
    Dim pRecZ
    Dim iSQL

    pFontName = "µ¸¿òÃ¼"
    pFontSize = 9
    pIEReOpenYN = "Y"
    pKeepSessionAliveTime = 1

    Set pRecZ = CreateObject("ADODB.Recordset")
    
    iSQL = "select upper(isnull(REFERENCE,'" & pFontName & "')) from b_configuration where MAJOR_CD = 'Z0052' and MINOR_CD = '1' and SEQ_NO = 1 "
    
    Call WriteToLog(iSQL)
    
    pRecZ.Open iSQL, gADODBConnString
            
    If Err.Number = 0 Then
       If pRecZ.EOF And pRecZ.BOF Then
       Else
          pFontName = pRecZ(0)
       End If
    End If
            
    If pRecZ.State = 1 Then
       pRecZ.Close
    End If
            
    iSQL = "select isnull(REFERENCE,'" & pFontSize & "') from b_configuration where MAJOR_CD = 'Z0052' and MINOR_CD = '2' and SEQ_NO = 1 "
     
    Call WriteToLog(iSQL)
    
    pRecZ.Open iSQL, gADODBConnString

    If Err.Number = 0 Then
       If pRecZ.EOF And pRecZ.BOF Then
       Else
          pFontSize = pRecZ(0)
       End If
    End If
            
    If pRecZ.State = 1 Then
       pRecZ.Close
    End If
            
    iSQL = "select isnull(REFERENCE,'" & pIEReOpenYN & "') from b_configuration where MAJOR_CD = 'Z0052' and MINOR_CD = '3' and SEQ_NO = 1 "
     
    Call WriteToLog(iSQL)
    
    pRecZ.Open iSQL, gADODBConnString

    If Err.Number = 0 Then
       If pRecZ.EOF And pRecZ.BOF Then
       Else
          pIEReOpenYN = pRecZ(0)
       End If
    End If
            
    If pRecZ.State = 1 Then
       pRecZ.Close
    End If
            
    iSQL = "select isnull(REFERENCE,'" & pKeepSessionAliveTime & "') from b_configuration where MAJOR_CD = 'Z0052' and MINOR_CD = '4' and SEQ_NO = 1 "
     
    Call WriteToLog(iSQL)
    
    pRecZ.Open iSQL, gADODBConnString

    If Err.Number = 0 Then
       If pRecZ.EOF And pRecZ.BOF Then
       Else
          pKeepSessionAliveTime = pRecZ(0)
       End If
    End If
    
    If CInt(pKeepSessionAliveTime) <= 0 Then
       pKeepSessionAliveTime = 1
    End If
            
    If pRecZ.State = 1 Then
       pRecZ.Close
    End If

    Set pRecZ = Nothing
            
End Sub


Sub GetuniConnectorInfo2(pW, pH)
   
    Dim pRecH
    Dim iSQL

    On Error Resume Next
    
    Set pRecH = CreateObject("ADODB.Recordset")
    
    iSQL = "select isnull(REFERENCE,'70') from b_configuration where MAJOR_CD = 'Z0053' and MINOR_CD = '1' and SEQ_NO = 1 "
    
    
    pRecH.Open iSQL, gADODBConnString
            
    If Err.Number = 0 Then
       If pRecH.EOF And pRecH.BOF Then
       Else
          pW = pRecH(0)
       End If
    Else
    End If
            
    If pRecH.State = 1 Then
       pRecH.Close
    End If
            
    iSQL = "select isnull(REFERENCE,'25') from b_configuration where MAJOR_CD = 'Z0053' and MINOR_CD = '2' and SEQ_NO = 1 "
    
    pRecH.Open iSQL, gADODBConnString

    If Err.Number = 0 Then
       If pRecH.EOF And pRecH.BOF Then
       Else
          pH = pRecH(0)
       End If
    Else
    End If
            
    If pRecH.State = 1 Then
       pRecH.Close
    End If
            
    Set pRecH = Nothing
            
End Sub





Function ExistsTable(ByVal pvTable)
    Dim adoRS
   
    ExistsTable = False

    Set adoRS = Server.CreateObject("ADODB.Recordset")
    adoRS.Open " SELECT OBJECT_ID('" & pvTable & "') ", gADODBConnString
   
    If Err.Number = 0 Then
       If adoRS.EOF And adoRS.BOF Then
       Else
          If IsNull(adoRS(0)) Then
          Else
             ExistsTable = True
          End If
       End If
    End If

    If adoRS.State = 1 Then
       adoRS.Close
    End If

    Set adoRs = Nothing

End Function



Sub WriteToLog(pLogData)

    On Error Resume Next
    Dim objFSO
    Dim objFile
    Dim pPath
    
    Exit Sub 
    
    pPath = "C:\ZConnectorAgent.txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    Set objFile = objFSO.OpenTextFile( pPath,8,True)
       
    objFile.WriteLine pLogData
   
    If Not (objFSO Is Nothing) Then
       Set objFSO = Nothing
    End If
    
    If Not (objFile Is Nothing) Then
       objFile.Close
       Set objFile = Nothing
    End If

End Sub


%>