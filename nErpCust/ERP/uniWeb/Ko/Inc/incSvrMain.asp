<%  %>
<!-- #Include file="CommResponse.inc" -->
<!-- #Include file="incSvrCcm.inc" -->
<!-- #Include file="incSvrVariables.inc" -->
<!-- #Include file="incSvrMessage.inc" -->


<Script Language=VBScript Runat=Server>

Sub LoadBasisGlobalInf()

	Dim iSepChar
	Dim xmlDoc
	Dim NodeNm

	On Error Resume Next 		    

	    
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = false 
		    
	xmlDoc.LoadXML(GetSessionStream)
	
	NodeNm = NodeNm1
	
		
	Call MappingCommon(xmlDoc,NodeNm)

	gADODBConnString     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gADODBConnString").text
	gDsnNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDsnNo").text
	gDBKind              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBKind").text
	gCanBeDebug          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCanBeDebug").text
	gISOLLVL             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gISOLLVL").text
	
	gPlant               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPlant").text
	gPlantNm             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPlantNm").text
	gSetupMod            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSetupMod").text
	gSeverity            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSeverity").text
	gColSep              = Chr(11)                    
	gRowSep			     = Chr(12)
    gLogoName            = Request.Cookies("unierp")("gLogoName")
    gLogo		         = Request.Cookies("unierp")("gLogo")    
    gADODBConnString     = gADODBConnString & " ; Workstation ID= " & gUsrId
    
	Set xmlDoc = Nothing

	iSepChar = "::"    
	gStrGlobalCollection = "2"                                                                             
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCanBeDebug                                   '0 : Debug Mode(1)
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(11)                                       '1
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(12)                                       '2
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(15)                                       '3
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(11)                                       '4
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(12)                                       '5
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & "1900-01-01"                                  '6
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & "YYYY-MM-DD"                                  '7
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & "-"                                           '8
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gADODBConnString                              '9
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gUsrId                                        '10
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gLang                                         '11
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCompany                                      '12
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gAPServer                                     '13
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gDBServer                                     '14
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gDatabase                                     '15
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Request.ServerVariables("REMOTE_ADDR")        '16
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCurrency                                     '17
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Request.ServerVariables("APPL_PHYSICAL_PATH") '18
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gISOLLVL                                      '19 Isolation Level(2)
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(20)                                       '20    
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & "72000"                                       '21 jslee 2003/04/26 
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(21)                                       '22 2004/08/12 unicode gAM
    gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCharSQLSet                                   '23 2004/08/12 unicode gAM  D :DBCS
	    
	gEnvInf =           gADODBConnString           & Chr(12)   '0
	gEnvInf = gEnvInf & gLang                      & Chr(12)   '1
	gEnvInf = gEnvInf & GetProgId()                & Chr(12)   '2
	gEnvInf = gEnvInf & gUsrId                     & Chr(12)   '3
	gEnvInf = gEnvInf & GetGlobalInf("gClientNm")  & Chr(12)   '4
	gEnvInf = gEnvInf & GetGlobalInf("gClientIp")  & Chr(12)   '5
	gEnvInf = gEnvInf & gUsrId                     & Chr(12)   '6
	gEnvInf = gEnvInf & gSeverity                  & Chr(12)   '7
	gEnvInf = gEnvInf & gDBKind                    & Chr(12)   '8
    
    gUDF6 = 0 
    gUDF7 = 0 
    gUDF8 = 0 
    gUDF9 = 0       
     
    If gCharSet = "U" Then
       adVarXChar = 202 ' adVarWChar
    Else
       adVarXChar = 200 ' adVarChar
    End If
   
    If gCharSet = "U" Then
       adXChar = 130  ' adWChar
    Else
       adXChar = 129  ' adChar
    End If
    
End sub

Sub MappingCommon(xmlDoc,NodeNm)

    gAPDateFormat        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateFormat").text
    gAPDateSeperator     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPDateSeperator").text
    gAPNum1000           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNum1000").text
    gAPNumDec            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPNumDec").text
    gAPServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAPServer").text
    gClientDateFormat    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateFormat").text  
    gClientDateSeperator = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientDateSeperator").text 
    gClientNum1000       = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNum1000").text   
    gClientNumDec        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNumDec").text        
    gComDateType         = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComDateType").text 
    gComNum1000          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNum1000").text
    gComNumDec           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gComNumDec").text           
    gConnectionString    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gConnectionString").text
    gDatabase            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDatabase").text
    gDateFormat          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormat").text
    gDateFormatYYYYMM    = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDateFormatYYYYMM").text
    gDBServer            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServer").text
    gLocRndPolicy        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLocRndPolicy").text    
    gTaxRndPolicy        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gTaxRndPolicy").text
    gAltNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gAltNo").text        
    gBConfMinorCD        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBConfMinorCD").text
    gCompany             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompany").text    
    gCompanyNm           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCompanyNm").text
    gCurrency            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCurrency").text
    gLang                = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLang").text
    gUsrId               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrId").text
    
End Sub

Function GetGlobalInf(ByVal pData)

	On error resume next 
	
    GetGlobalInf = GetGlobalInf2(NodeNm2,pData)

End Function


Function GetGlobalInf2(ByVal pNodeName,ByVal pData)

	Dim xmlDoc
	
	On Error Resume Next 
	
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = False 
    xmlDoc.LoadXML (GetSessionStream)
	
	GetGlobalInf2	= xmlDoc.selectSingleNode("/uniERP/" & pNodeName & "/" & pData ).text   

	Set xmlDoc = Nothing

End Function


Function GetGlobalInf3(ByVal pXML ,ByVal pNodeName,ByVal pData)


    Dim xmlDOMDocumentX

    Set xmlDOMDocumentX = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDOMDocumentX.async = False 
	    
	xmlDOMDocumentX.loadXML(pXML)
	
	GetGlobalInf3	= xmlDOMDocumentX.selectSingleNode("/uniERP/" & pNodeName & "/" & pData ).text   

	Set xmlDOMDocumentX = Nothing

End Function


Function GetGlobalData(pData)   '2003-08-07 leejinsoo
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


'==========================================================================================
' Name : GetProgId
' Desc : Get current program id 
'==========================================================================================
Function GetProgId()

	Dim strLoc, iPos , iLoc, strAspName
	
	strLoc = Request.ServerVariables("URL")
	
	iLoc = 1: iPos = 0
	
	Do Until iLoc <= 0						
		iLoc = inStr(iPos+1, strLoc, "/")
		If iLoc <> 0 Then iPos = iLoc
	Loop
		
	strAspName = Right(strLoc, Len(strLoc) - iPos)
	GetProgId = Left(strAspName, Len(strAspName) - Len(".ASP"))	
	
End Function

'==============================================================================
' Hide Current Window
'==============================================================================
Sub HideStatusWnd()
	Response.Write "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
	Response.Write "Sub Document_onReadyStateChange()" & vbCrLf
	Response.Write " On Error Resume Next "            & vbCrLf
	Response.Write "Call parent.BtnDisabled(False)"    & vbCrLf	
	Response.Write "Call parent.LayerShowHide(0)"      & vbCrLf
	Response.Write "Call parent.RestoreToolBar()"      & vbCrLf
	Response.Write "End Sub"  & vbCrLf
	Response.Write "</" & "Script" & ">" & vbCrLf
End Sub

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

   
'========================================================================================
' Function Name : ConvSPChars
' Function Desc : Replace " with ""
'========================================================================================
Function ConvSPChars(strVal)
	ConvSPChars = Replace("" & strVal, """", """""")
End Function 


'=============================================================================
' Function Name : LoadTab
' Function Desc : LoadTab
'=============================================================================

Function LoadTab(objTarget, iTabNo, iLoc)
    Dim strHTML
    
    If iTabNo > 0 Then
        If iLoc = I_INSCRIPT Then
    		strHTML = "Call parent.ClickTab" & iTabNo
    		Response.Write strHTML
    	ElseIf iLoc = I_MKSCRIPT Then
    		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
    		strHTML = strHTML & "Call parent.ClickTab" & iTabNo & vbCrLf
    		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
    		Response.Write strHTML
    	End If
	End If

	Call HTMLFocus(objTarget, iLoc)    
	
End Function

'======================================================================================================
' 설명 : HTMLFocus
' 기능 : make relevant object focused on the Client Side 
'======================================================================================================
Function HTMLFocus(objTarget,  iLoc)
    Dim strHTML
    
	If iLoc = I_INSCRIPT Then
		strHTML = strHTML & objTarget & ".focus" & vbCrLf
		strHTML = strHTML & objTarget & ".select" & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
   	    strHTML = strHTML & " On Error Resume Next "  & vbCrLf
		strHTML = strHTML & objTarget & ".focus" & vbCrLf
		strHTML = strHTML & objTarget & ".select" & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

</Script>

