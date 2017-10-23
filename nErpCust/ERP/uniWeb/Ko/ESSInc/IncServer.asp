<% Response.Buffer = True %>
<!-- #Include file="CommResponse.inc" -->
<!-- #Include file="incSvrVariables.inc" -->
<!-- #Include file="incSvrCcm.inc" -->
<!-- #Include file="incSvrDate.inc" -->
<!-- #Include file="incSvrDBAgent.inc" -->
<!-- #Include file="incSvrNumber.inc" -->
<!-- #Include file="incSvrMessage.inc" -->
<!-- #Include file="incSvrDBAgentVariables.inc" -->
<%

Call GetGlobalVar

%>

<Script Language=VBScript Runat=Server>

'==============================================================================
'
'==============================================================================
Sub GetGlobalVar
	
	gLogoName	= Request.Cookies("unierp")("gLogoName")
    gLogo		= Request.Cookies("unierp")("gLogo")   
    gColSep		= Chr(11)                    
	gRowSep		= Chr(12)
	
    If Request.Cookies("unierp")("gVersion") = "V2.5" Then 
		Call GetGlobalVarFromCookies
	Else
        If Request.Cookies("unierp")("gAPPKind") = "ESS" Then 
		   Call GetESSGlobalVarFromXML
		Else
		   Call GetGlobalVarFromXML
		End If 
	End If 

    Call LoadNumericData()
	Call LoadDataForMessage()     		
    Call LoadDataForVB() 
    
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

Sub GetGlobalVarFromXML
	
	Dim xmlDoc
	Dim NodeNm
			
    On Error Resume Next 

	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = False 
	    
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
	gUserIdKind          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUserIdKind").text
	    
	NodeNm = NodeNm2
	
	gDBLoginPwd          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginPwd").text
	gBizArea             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBizArea").text
	gBizUnit             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gBizUnit").text
	gChangeOrgId         = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gChangeOrgId").text
	gCostCenter          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCostCd").text
	gCountry             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCountry").text
	gDepart              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDepart").text
	gDBLoginID           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginID").text
	gDBServerIP          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerIP").text		
	gDBServerNm          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerNm").text	    
	gFiscCnt             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscCnt").text
	gFiscEnd             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscEnd").text
	gFiscStart           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gFiscStart").text
	gIntDeptCd           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gIntDeptCd").text
	gIm_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gIm_Post_Flag").text
	gLoginDt             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLoginDt").text    
	gLogonGp             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gLogonGp").text
	gPurOrg              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPurOrg").text
	gPurGrp              = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPurGrp").text
	gSalesGrp            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSalesGrp").text
	gSalesOrg            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSalesOrg").text
	gSo_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gSo_Post_Flag").text
	gPo_Post_Flag        = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gPo_Post_Flag").text
	gStorageLoc          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gStorageLoc").text
	gUsrEngName          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrEngName").text
	gUsrNm               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrNm").text
	gWorkCenter          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gWorkCenter").text
	gClientIp            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientIp").text
	gClientNm            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gClientNm").text
	
	NodeNm = NodeNm3
	gEWare				 = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gEWareURL").text  

	Set xmlDoc = Nothing	                            
End Sub

Sub GetESSGlobalVarFromXML
    Dim xmlDoc
    Dim NodeNm

    On Error Resume Next 

    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")      
    
    
    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = False 
	xmlDoc.LoadXML(GetSessionStream)	

    NodeNm = NodeNm1
    
	gADODBConnString     = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gADODBConnString").text
	gDsnNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDsnNo").text
	gCanBeDebug          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gCanBeDebug").text
	gISOLLVL             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gISOLLVL").text
	
    Call MappingCommon(xmlDoc,NodeNm)
   
    NodeNm = NodeNm2

    gDBLoginID           = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginID").text
	gDBLoginPwd          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBLoginPwd").text
    gDBServerIP          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerIP").text      
    gDBServerNm          = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDBServerNm").text      
    gUsrNm               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gUsrNm").text
    gEmpNo               = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gEmpNo").text
    gProAuth             = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gProAuth").text
    gDeptAuth            = xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "gDeptAuth ").text           

    Set xmlDoc = Nothing                                
End Sub


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


Sub GetGlobalVarFromCookies

	' LoadBasisGlobalInf
    gADODBConnString     = Request.Cookies("unierp")("gADODBConnString")    
    gAPDateFormat        = Request.Cookies("unierp")("gAPDateFormat")
    gAPDateSeperator     = Request.Cookies("unierp")("gAPDateSeperator")
    gAPNum1000           = Request.Cookies("unierp")("gAPNum1000")
    gAPNumDec            = Request.Cookies("unierp")("gAPNumDec")    
    gAPServer            = Request.Cookies("unierp")("gAPServer")
    gClientDateFormat    = Request.Cookies("unierp")("gClientDateFormat")
    gClientDateSeperator = Request.Cookies("unierp")("gClientDateSeperator")    
    gClientNum1000       = Request.Cookies("unierp")("gClientNum1000")
    gClientNumDec        = Request.Cookies("unierp")("gClientNumDec")
	gComDateType		 = Request.Cookies("unierp")("gComDateType")
    gComNum1000			 = Request.Cookies("unierp")("gComNum1000")  
    gComNumDec           = Request.Cookies("unierp")("gNumDec")
    gConnectionString    = Request.Cookies("unierp")("gConnectionString")
    gDatabase            = Request.Cookies("unierp")("gDatabase")
    gDateFormat          = Request.Cookies("unierp")("gDateFormat")
    gDateFormatYYYYMM	 = Request.Cookies("unierp")("gDateFormatYYYYMM")
    gDBServer            = Request.Cookies("unierp")("gDBServer")
    gDsnNo               = Request.Cookies("unierp")("gDsnNo")    
    gLocRndPolicy        = Request.Cookies("unierp")("gLocRndPolicy")	
    gTaxRndPolicy        = Request.Cookies("unierp")("gTaxRndPolicy")	
    gAltNo               = Request.Cookies("unierp")("gAltNo")    
	gBConfMinorCD        = Request.Cookies("unierp")("gBConfMinorCD")	
	gCompany             = Request.Cookies("unierp")("gCompany")
    gCompanyNm           = Request.Cookies("unierp")("gCompanyNm")
    gCurrency            = Request.Cookies("unierp")("gCurrency")
    gLang                = Request.Cookies("unierp")("gLang")
    gPlant               = Request.Cookies("unierp")("gPlant")
    gPlantNm             = Request.Cookies("unierp")("gPlantNm") 
    gSetupMod            = Request.Cookies("unierp")("gSetupMod")
    gSeverity            = Request.Cookies("unierp")("gSeverity")     
    gUsrId               = Request.Cookies("unierp")("gUsrId")
    
    ' GetGlobalInf
    gBizArea             = Request.Cookies("unierp")("gBizArea")
    gBizUnit             = Request.Cookies("unierp")("gBizUnit") 
    gChangeOrgId         = Request.Cookies("unierp")("gChangeOrgId")        
    gCostCenter          = Request.Cookies("unierp")("gCostCd")
    gCountry             = Request.Cookies("unierp")("gCountry")    
    gDepart              = Request.Cookies("unierp")("gDepartment")    
    gDBLoginID           = Request.Cookies("unierp")("gDBLoginID")
    gDBLoginPwd          = Request.Cookies("unierp")("gDBLoginPwd")    
    gDBServerIP          = Request.Cookies("unierp")("gDBServerIP")		
    gDBServerNm          = Request.Cookies("unierp")("gDBServerNm")
    gFiscCnt             = Request.Cookies("unierp")("gFiscCnt")
    gFiscEnd             = Request.Cookies("unierp")("gFiscEnd")
    gFiscStart           = Request.Cookies("unierp")("gFiscStart")
    gIntDeptCd           = Request.Cookies("unierp")("gIntDeptCd")
    gIm_Post_Flag        = Request.Cookies("unierp")("gPostFlagIm")
    gLoginDt             = Request.Cookies("unierp")("gLoginDt")
    gLogonGp             = Request.Cookies("unierp")("gLogonGp")
    gPurOrg              = Request.Cookies("unierp")("gPurOrg")
    gPurGrp              = Request.Cookies("unierp")("gPurGrp")
    gSalesGrp            = Request.Cookies("unierp")("gSalesGrp")
    gSalesOrg            = Request.Cookies("unierp")("gSalesOrg")    
    gSo_Post_Flag        = Request.Cookies("unierp")("gPostFlagSo")
    gPo_Post_Flag        = Request.Cookies("unierp")("gPostFlagPo")
    gStorageLoc          = Request.Cookies("unierp")("gStorageLoc")
    gUsrEngName          = Request.Cookies("unierp")("gUsrEngNm")    
    gUsrNm               = Request.Cookies("unierp")("gUsrNm")
    gWorkCenter          = Request.Cookies("unierp")("gWorkCenter")        
    gClientIp            = Request.Cookies("unierp")("gClientIp")
    gClientNm            = Request.Cookies("unierp")("gClientNm")
    gEWare				 = Request.Cookies("unierp")("gEWareURL")     

End sub


'==========================================================================================
'
'==========================================================================================
Sub LoadDataForMessage()

    gEnvInf =           gADODBConnString  & gRowSep   '0
    gEnvInf = gEnvInf & gLang             & gRowSep   '1
    gEnvInf = gEnvInf & GetProgId()	      & gRowSep   '2
    gEnvInf = gEnvInf & gUsrId            & gRowSep   '3
    gEnvInf = gEnvInf & gClientNm         & gRowSep   '4
    gEnvInf = gEnvInf & gClientIp         & gRowSep   '5
    gEnvInf = gEnvInf & gUsrId            & gRowSep   '6
    gEnvInf = gEnvInf & gSeverity         & gRowSep   '7
    gEnvInf = gEnvInf & gDBKind           & gRowSep   '8

End Sub
'==========================================================================================
'
'==========================================================================================
Sub LoadDataForVB()

    Dim iSepChar

	iSepChar = "::"    
	gStrGlobalCollection = "2" 
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCanBeDebug                                   '0 Debug Mode(1)
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gColSep                                       '1
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gRowSep                                       '2
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(15)                                       '3
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gColSep                                       '4
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gRowSep                                       '5
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
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & gISOLLVL           '19 Isolation level(2)
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(20)                                       '20  gBM
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & "300"                                         '21 2003/04/26       
	gStrGlobalCollection = gStrGlobalCollection & iSepChar & Chr(21)                                       '22 2004/08/12 unicode gAM
    gStrGlobalCollection = gStrGlobalCollection & iSepChar & gCharSQLSet                                   '23 2004/08/12 unicode gAM  D :DBCS
	
End Sub

'==========================================================================================
'
'==========================================================================================
Sub LoadNumericData()

    Set ggAmtOfMoney = New TB19029
    Set ggQty        = New TB19029
    Set ggUnitCost   = New TB19029
    Set ggExchRate   = New TB19029
          
    ggAmtOfMoney.DecPoint  = Request.Cookies("unierp")("gAmtOfMoney")
    ggQty.DecPoint         = Request.Cookies("unierp")("gQty")
    ggUnitCost.DecPoint    = Request.Cookies("unierp")("gUnitCost")
    ggExchRate.DecPoint    = Request.Cookies("unierp")("gExchRate")
    
    ggAmtOfMoney.RndPolicy = Request.Cookies("unierp")("gAmtOfMoneyRndPolicy")
    ggQty.RndPolicy        = Request.Cookies("unierp")("gQtyRndPolicy")
    ggUnitCost.RndPolicy   = Request.Cookies("unierp")("gUnitCostRndPolicy")
    ggExchRate.RndPolicy   = Request.Cookies("unierp")("gExchRateRndPolicy")
    
    ggAmtOfMoney.RndUnit   = Request.Cookies("unierp")("gAmtOfMoneyRndUnit")
    ggQty.RndUnit          = Request.Cookies("unierp")("gQtyRndUnit")
    ggUnitCost.RndUnit     = Request.Cookies("unierp")("gUnitCostRndUnit")
    ggExchRate.RndUnit     = Request.Cookies("unierp")("gExchRateRndUnit")    
    
    gBDataType             = Request.Cookies("unierp")("gBDataType")         '0
    gBCurrency             = Request.Cookies("unierp")("gBCurrency")         '1
    gBDecimals             = Request.Cookies("unierp")("gBDecimals")         '2
    gBRoundingUnit         = Request.Cookies("unierp")("gBRoundingUnit")     '3
    gBRoundingPolicy       = Request.Cookies("unierp")("gBRoundingPolicy")   '4
    
    gBDataType             = Split(gBDataType      ,gColSep) 
    gBCurrency             = Split(gBCurrency      ,gColSep) 
    gBDecimals             = Split(gBDecimals      ,gColSep) 
    gBRoundingUnit         = Split(gBRoundingUnit  ,gColSep) 
    gBRoundingPolicy       = Split(gBRoundingPolicy,gColSep) 
    
End Sub

Function GetGlobalInf3(ByVal pXML ,ByVal pNodeName,ByVal pData)


    Dim xmlDOMDocumentX

    Set xmlDOMDocumentX = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDOMDocumentX.async = False 
	    
	xmlDOMDocumentX.loadXML(pXML)
	
	GetGlobalInf3	= xmlDOMDocumentX.selectSingleNode("/uniERP/" & pNodeName & "/" & pData ).text   

	Set xmlDOMDocumentX = Nothing

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
' Function Desc : replace " with ""
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
' Function Name : HTMLFocus
' Function Desc : make relevant object focused on the Client Side 
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

'-------------------------------------------------------------------------------------------------------------------------------
' Sub Name   : 
'-------------------------------------------------------------------------------------------------------------------------------
Sub SubFillRemBodyTD5656(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD> </TR>"
   Next    
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Sub Name   : 
'-------------------------------------------------------------------------------------------------------------------------------
Sub SubFillRemBodyTD656(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD656 NOWRAP>&nbsp;</TD></TR>"
   Next    
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Sub Name   : 
'-------------------------------------------------------------------------------------------------------------------------------
Sub SubFillRemBodyTD56(iTer)
   Dim iDx 
   For iDx = 0 To iTer
       Response.Write "<TR> <TD CLASS=TD5 NOWRAP>&nbsp;</TD> <TD CLASS=TD6 NOWRAP>&nbsp;</TD></TR>"
   Next    
End Sub

</Script>
<Script Language=VBScript >
    Dim gServerIP
    Dim gLogoName
    Dim gLogo
    Dim gFontName
    Dim gFontSize
    Dim gCharSet
    Dim gCharSQLSet
    
    If Instr(document.location.href,"http://") > 0 Then
       gServerIP = "http://<%= request.servervariables("server_name") %>"
    Else
       gServerIP = "https://<%= request.servervariables("server_name") %>"
    End If   

    gLogoName = "<%=Request.Cookies("unierp")("gLogoName")%>"
    gLogo     = "<%=Request.Cookies("unierp")("gLogo")%>"
    gFontName = "<%=Request.Cookies("unierp")("gFontName")%>"
    gFontSize = "<%=Request.Cookies("unierp")("gFontSize")%>"    

    gCharSet    = "<%=Request.Cookies("unierp")("gCharSet")%>"
    gCharSQLSet = "<%=Request.Cookies("unierp")("gCharSQLSet")%>"
</Script>

