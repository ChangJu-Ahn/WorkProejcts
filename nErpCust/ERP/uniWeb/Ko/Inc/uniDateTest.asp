<% Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="CommResponse.inc" -->
<!-- #Include file="incSvrVariables.inc" -->
<!-- #Include file="incSvrCcm.inc" -->
<!-- #Include file="incSvrDate.inc" -->
<!-- #Include file="incSvrDBAgent.inc" -->
<!-- #Include file="incSvrNumber.inc" -->
<!-- #Include file="incSvrMessage.inc" -->
<!-- #Include file="incSvrString.inc" -->
<!-- #Include file="SError.inc" -->
<%
Call GetGlobalVar
   Response.Write "<BR><PRE>gClientDateFormat => [" & gClientDateSeperator & "][" & gClientDateFormat & "]"
   Response.Write "<BR><PRE>gAPDateFormat     => [" & gAPDateSeperator     & "][" & gAPDateFormat     & "]"
   Response.Write "<BR><PRE>gDateFormat       => [" & gComDateType         & "][" & gDateFormat       & "]"
   Response.Write "<BR><PRE>gDateFormatYYYYMM => [" & gComDateType         & "][" & gDateFormatYYYYMM & "]"
   Response.Write "<BR><PRE>gServerDateFormat => [" & gServerDateType      & "][" & gServerDateFormat & "]"
   Response.Write "<BR><PRE>"

   BaseDate    = GetSvrDate                                                                  'Get DB Server Date
   CBaseDate   = Date
   CPD         = UNIConvDateAToB(BaseDate ,gServerDateFormat,gDateFormat)
   APD         = UNIConvDateAToB(BaseDate ,gServerDateFormat,gAPDateFormat)

   Response.Write "<BR><PRE>[" & GetSvrDateTime & "] DB" :
   Response.Write "<BR><PRE>[" & BaseDate       & "] DB"
   Response.Write "<BR><PRE>[" & CBaseDate      & "] AP"
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>UNIGetLastDay [" & UNIGetLastDay (BaseDate,gServerDateFormat) & "] Last  Day"
   Response.Write "<BR><PRE>UNIGetFirstDay[" & UNIGetFirstDay(BaseDate,gServerDateFormat) & "] First Day"
   Response.Write "<BR><PRE>UNIDateAdd    [" & UniDateAdd("m", -2, BaseDate,gServerDateFormat) & "] Month - 2 "
   Response.Write "<BR><PRE>UNIDateAdd    [" & UniDateAdd("m",  2, BaseDate,gServerDateFormat) & "] Month + 2 "
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>UNIConvDateAToB        : gServerDateFormat [" & gServerDateFormat & "][" & BaseDate & "]-> gDateFormat      [" & gDateFormat       & "][" & UNIConvDateAToB(BaseDate ,gServerDateFormat,gDateFormat) & "]"
   Response.Write "<BR><PRE>UNIConvDateAToB        : gDateFormat       [" & gDateFormat       & "][" & CPD      & "]-> gServerDateFormat[" & gServerDateFormat & "][" & UNIConvDateAToB(CPD      ,gDateFormat,gServerDateFormat) & "]"
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>uniConvDate            : gDateFormat       [" & gDateFormat       & "][" & CPD      & "]-> gServerDateFormat[" & gServerDateFormat & "][" & uniConvDate(CPD) & "]"
   Response.Write "<BR><PRE>UNIConvDateCompanyToDB : gAPDateFormat     [" & gDateFormat       & "][" & CPD      & "]-> gServerDateFormat[" & gServerDateFormat & "][" & UNIConvDateCompanyToDB(CPD,null) & "]"
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>UNIDateClientFormat    : gAPDateFormat     [" & gAPDateFormat     & "][" & APD      & "]-> gDateFormat      [" & gDateFormat       & "][" & UNIDateClientFormat(APD) & "]"
   Response.Write "<BR><PRE>UNIConvDateDBToCompany : gAPDateFormat     [" & gAPDateFormat     & "][" & APD      & "]-> gDateFormat      [" & gDateFormat       & "][" & UNIConvDateDBToCompany(APD,null) & "]"
   Response.Write "<BR><PRE>UNIMonthClientFormat   : gAPDateFormat     [" & gAPDateFormat     & "][" & APD      & "]-> gDateFormat      [" & gDateFormat       & "][" & UNIMonthClientFormat(APD) & "]"
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>UniConvDateToYYYYMMDD  : gDateFormat       [" & gDateFormat       & "][" & CPD      & "]->[-][" & UniConvDateToYYYYMMDD(CPD,gDateFormat      ,"-") & "]"
   Response.Write "<BR><PRE>UniConvDateToYYYYMMDD  : gAPDateFormat     [" & gAPDateFormat     & "][" & APD      & "]->[/][" & UniConvDateToYYYYMMDD(APD,gAPDateFormat    ,"/") & "]"
   Response.Write "<BR><PRE>UniConvDateToYYYYMMDD  : gAPDateFormat     [" & gAPDateFormat     & "][" & APD      & "]->[/][" & UniConvDateToYYYYMMDD(APD,gAPDateFormat    ,"/") & "]"
   Response.Write "<BR><PRE>"
   Response.Write "<BR><PRE>UniConvYYYYMMDDToDate  : [2001,04,06]->[gDateFormat      ][" & gDateFormat       & "][" & UniConvYYYYMMDDToDate(gDateFormat,"2001","04","06") & "]"
   Response.Write "<BR><PRE>UniConvYYYYMMDDToDate  : [2001,04,06]->[gAPDateFormat    ][" & gAPDateFormat     & "][" & UniConvYYYYMMDDToDate(gAPDateFormat,"2001","04","06") & "]"
   Response.Write "<BR><PRE>UniConvYYYYMMDDToDate  : [2001,04,06]->[gClientDateFormat][" & gClientDateFormat & "][" & UniConvYYYYMMDDToDate(gClientDateFormat,"2001","04","06") & "]"
   Response.Write "<BR><PRE>UniConvYYYYMMDDToDate  : [2001,04,06]->[gServerDateFormat][" & gServerDateFormat & "][" & UniConvYYYYMMDDToDate(gServerDateFormat,"2001","04","06") & "]"
%>





<Script Language=VBScript Runat=Server>

'==============================================================================
'
'==============================================================================
Sub GetGlobalVar

    Dim strTemp

   'Set ggAmtExOfMoney = New TB19029           'For Hermes
    Set ggAmtOfMoney = New TB19029
    Set ggQty        = New TB19029
    Set ggUnitCost   = New TB19029
    Set ggExchRate   = New TB19029
    
    
'   gA1015               = Request.Cookies("unierp")("g1015")   'For Hermes
'   gA1053               = Request.Cookies("unierp")("g1053")   'For Hermes
'   gA1057               = Request.Cookies("unierp")("g1057")   'For Hermes
	gADODBConnString     = Request.Cookies("unierp")("gADODBConnString")
    gAltNo               = Request.Cookies("unierp")("gAltNo")    
    gAPDateFormat        = Request.Cookies("unierp")("gAPDateFormat")
    gAPDateSeperator     = Request.Cookies("unierp")("gAPDateSeperator")
    gAPNum1000           = Request.Cookies("unierp")("gAPNum1000")
    gAPNumDec            = Request.Cookies("unierp")("gAPNumDec")
    gAPServer            = Request.Cookies("unierp")("gAPServer")
    gBizArea             = Request.Cookies("unierp")("gBizArea")
    gBizUnit             = Request.Cookies("unierp")("gBizUnit") 
'   gBizUnitNm           = Request.Cookies("unierp")("gBizUnitNm")   'For Hermes
    gChangeOrgId         = Request.Cookies("unierp")("gChangeOrgId")
    gClientDateFormat    = Request.Cookies("unierp")("gClientDateFormat")
    gClientDateSeperator = Request.Cookies("unierp")("gClientDateSeperator")
    gClientIp            = Request.ServerVariables("REMOTE_ADDR")
    gClientNm            = Request.ServerVariables("REMOTE_ADDR")
    gClientNum1000       = Request.Cookies("unierp")("gClientNum1000")
    gClientNumDec        = Request.Cookies("unierp")("gClientNumDec")
    gComNumDec           = Request.Cookies("unierp")("gNumDec")
    gCompany             = Request.Cookies("unierp")("gCompany")
    gCompanyNm           = Request.Cookies("unierp")("gCompanyNm")
    gConnectionString    = Request.Cookies("unierp")("gConnectionString")
    gCostCenter          = Request.Cookies("unierp")("gCostCd")
'   gCostCenterNm        = Request.Cookies("unierp")("gCostNm")      'For Hermes
    gCountry             = Request.Cookies("unierp")("gCountry")
    gCurrency            = Request.Cookies("unierp")("gCurrency")
    gDatabase            = Request.Cookies("unierp")("gDatabase")
    gDateFormat          = Request.Cookies("unierp")("gDateFormat")
    gDepart              = Request.Cookies("unierp")("gDepartment")
    gDBLoginID           = Request.Cookies("unierp")("gDBLoginID")
    gDBLoginPwd          = Request.Cookies("unierp")("gDBLoginPwd")
    gDBServer            = Request.Cookies("unierp")("gDBServer")
    gDBServerIP          = Request.Cookies("unierp")("gDBServerIP")
    gDBServerNm          = Request.Cookies("unierp")("gDBServerNm")
    gDsnNo               = Request.Cookies("unierp")("gDsnNo")    
    gFiscCnt             = Request.Cookies("unierp")("gFiscCnt")
    gFiscEnd             = Request.Cookies("unierp")("gFiscEnd")
    gFiscStart           = Request.Cookies("unierp")("gFiscStart")
    gIntDeptCd           = Request.Cookies("unierp")("gIntDeptCd")
    gIm_Post_Flag        = Request.Cookies("unierp")("gPostFlagIm")
    gLang                = Request.Cookies("unierp")("gLang")
    gLocRndPolicy        = Request.Cookies("unierp")("gLocRndPolicy")
    gLoginDt             = Request.Cookies("unierp")("gLoginDt")
    gLogonGp             = Request.Cookies("unierp")("gLogonGp")
    gPlant               = Request.Cookies("unierp")("gPlant")
    gPlantNm             = Request.Cookies("unierp")("gPlantNm")      
    gProgId              = GetProgId()
    gPurOrg              = Request.Cookies("unierp")("gPurOrg")
    gPurGrp              = Request.Cookies("unierp")("gPurGrp")
    gSalesGrp            = Request.Cookies("unierp")("gSalesGrp")
    gSalesOrg            = Request.Cookies("unierp")("gSalesOrg")
    gSetupMod            = Request.Cookies("unierp")("gSetupMod")
    gSeverity            = Request.Cookies("unierp")("gSeverity")
    gSo_Post_Flag        = Request.Cookies("unierp")("gPostFlagSo")
    gPo_Post_Flag        = Request.Cookies("unierp")("gPostFlagPo")
    gStorageLoc          = Request.Cookies("unierp")("gStorageLoc")
    gTaxRndPolicy        = Request.Cookies("unierp")("gTaxRndPolicy")
'   gToExcelYesNo        = Request.Cookies("unierp")("gToExcelYesNo") ' For Hermes
   'gUsrAuthLvl          = GetUsrAuthLvl()
    gUsrEngName          = Request.Cookies("unierp")("gUsrEngNm")
    gUsrId               = Request.Cookies("unierp")("gUsrId")
    gUsrNm               = Request.Cookies("unierp")("gUsrNm")
    gWorkCenter          = Request.Cookies("unierp")("gWorkCenter")

    If gComNumDec  = "." Then
       gComNum1000 = ","
    Elseif gComNumDec = "," Then
       gComNum1000 = "."
    End if

    strTemp = gDateFormat
    strTemp = Replace(strTemp, "Y","")
    strTemp = Replace(strTemp, "M","")
    strTemp = Replace(strTemp, "D","")

    gComDateType = Left(strTemp,1)

    gDateFormatYYYYMM      = Replace(gDateFormat      ,"DD" & gComDateType ,"")
    gDateFormatYYYYMM      = Replace(gDateFormatYYYYMM,gComDateType & "DD" ,"")

   'ggAmtExOfMoney.DecPoint = Request.Cookies("unierp")("gAmtExOfMoney")          ' For Hermes
    ggAmtOfMoney.DecPoint   = Request.Cookies("unierp")("gAmtOfMoney")
    ggQty.DecPoint          = Request.Cookies("unierp")("gQty")
    ggUnitCost.DecPoint     = Request.Cookies("unierp")("gUnitCost")
    ggExchRate.DecPoint     = Request.Cookies("unierp")("gExchRate")
    
   'ggAmtExOfMoney.RndPolicy = Request.Cookies("unierp")("gAmtExOfMoneyRndPolicy")          ' For Hermes
    ggAmtOfMoney.RndPolicy = Request.Cookies("unierp")("gAmtOfMoneyRndPolicy")
    ggQty.RndPolicy        = Request.Cookies("unierp")("gQtyRndPolicy")
    ggUnitCost.RndPolicy   = Request.Cookies("unierp")("gUnitCostRndPolicy")
    ggExchRate.RndPolicy   = Request.Cookies("unierp")("gExchRateRndPolicy")
    
   'ggAmtExOfMoney.RndUnit   = Request.Cookies("unierp")("gAmtExOfMoneyRndUnit")          ' For Hermes
    ggAmtOfMoney.RndUnit   = Request.Cookies("unierp")("gAmtOfMoneyRndUnit")
    ggQty.RndUnit          = Request.Cookies("unierp")("gQtyRndUnit")
    ggUnitCost.RndUnit     = Request.Cookies("unierp")("gUnitCostRndUnit")
    ggExchRate.RndUnit     = Request.Cookies("unierp")("gExchRateRndUnit")    
    
    gBDataType             = Request.Cookies("unierp")("gBDataType")         '0
    gBCurrency             = Request.Cookies("unierp")("gBCurrency")         '1
    gBDecimals             = Request.Cookies("unierp")("gBDecimals")         '2
    gBRoundingUnit         = Request.Cookies("unierp")("gBRoundingUnit") '3
    gBRoundingPolicy       = Request.Cookies("unierp")("gBRoundingPolicy")   '4
    
    gBDataType             = Split(gBDataType      ,Chr(11)) 
    gBCurrency             = Split(gBCurrency      ,Chr(11)) 
    gBDecimals             = Split(gBDecimals      ,Chr(11)) 
    gBRoundingUnit         = Split(gBRoundingUnit  ,Chr(11)) 
    gBRoundingPolicy       = Split(gBRoundingPolicy,Chr(11)) 

    gEnvInf =           gConnectionString  &  Chr(12)   '0
    gEnvInf = gEnvInf & gLang              &  Chr(12)   '1
    gEnvInf = gEnvInf & gProgId            &  Chr(12)   '2
    gEnvInf = gEnvInf & gUsrId             &  Chr(12)   '3
    gEnvInf = gEnvInf & gClientNm          &  Chr(12)   '4
    gEnvInf = gEnvInf & gClientIp          &  Chr(12)   '5
    gEnvInf = gEnvInf & gUsrId             &  Chr(12)   '6
    gEnvInf = gEnvInf & gSeverity          &  Chr(12)   '7
    
    gLogoName = Request.Cookies("unierp")("gLogoName")
    
End sub

'==========================================================================================
' Name : GetProgId
' Desc : Get current program id 
'==========================================================================================
Function GetProgId()

	GetProgId = Request("strRequestMenuID")
	
End Function

Function GetUsrAuthLvl()
 
    Dim za0015
    Dim GroupCount
    Dim LngRow
    Dim strData
    
    Set za0015 = Server.CreateObject("za0015.za0015CheckUsrAuthLvl")
    
    If Err.Number <> 0 Then
		Set za0015 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If
	
	za0015.ImportPgmZCoMastMnuMnuId = GetProgId 
	za0015.ImportUsrZUsrMastRecUsrId = gUsrId
	
	za0015.ComCfg = gConnectionString
	za0015.Execute	
	
	GroupCount = za0015.ExportGroupCount

    strData = strData & gUsrId & space(13 - len(gUsrId)) & _
              GetProgId & space(15 - len(GetProgId))
              
    For LngRow = 1 To GroupCount
       strData = strData & za0015.ExportAuthZUsrPgmRecordSetOrgType(LngRow)
       strData = strData & za0015.ExportAuthZUsrPgmRecordSetRAuth(LngRow)
       strData = strData & za0015.ExportAuthZUsrPgmRecordSetCudAuth(LngRow)
       strData = strData & za0015.ExportOrgZUsrOrgMastOrgCd(LngRow) & _
                 space(10 - len(za0015.ExportOrgZUsrOrgMastOrgCd(LngRow))) 
    Next
    
    strData = strData & "/"
    
    If len(strData) = 29 Then
       GetUsrAuthLvl = ""              'if Authority Value does not Exist, return  null string 
    Else   
	   GetUsrAuthLvl = strData
	End if
	
	Set za0015 = nothing   
	
End Function 

'=============================================================================
' Function Name  : LoadTab
' Parameter      : strString -> Message text
'                  strTarget -> "%"
' Description    : This function is counting "%" value
' Return Value   : "%" Count
'=============================================================================

Function LoadTab(objTarget, iTabNo, iLoc)
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
'   Dim gFileServerIP              'For Hermes

'   gFileServerIP = "70.2.226.100" 'For Hermes
    
    If Instr(document.location.href,"http://") > 0 Then
       gServerIP = "http://<%= request.servervariables("server_name") %>"
    Else
       gServerIP = "https://<%= request.servervariables("server_name") %>"
    End If   

    gLogoName = "<%=Request.Cookies("unierp")("gLogoName")%>"
</Script>
<%
   Dim pDebug
   pDebug = Server.MapPath(Request.ServerVariables("PATH_INFO"))
   If Mid(UCase(pDebug),InStr(UCase(pDebug),".ASP") -2 ,1) = "B" Then
      Response.Write "<BODY style=""background-color:lightsteelblue;font-family:verdana,arial""></Body>"
   End If
%>