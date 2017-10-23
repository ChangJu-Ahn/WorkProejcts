<%							

'=======================================================================================================
'
'=======================================================================================================
Sub LoadInfTB19029(ByVal pCurrency, ByVal pFormType, ByVal pModuleCd)
    
    On Error Resume Next												            '☜: Server Side Process
    
    If Trim(pCurrency) = "" Then
       pCurrency = Read_B_Company(gADODBConnString)
    End If
    Call QueryFormat(pCurrency, pFormType, pModuleCd,"COOKIE","",gADODBConnString)    
    
End Sub    

'=======================================================================================================
'
'=======================================================================================================
Sub LoadInfTB19029A(ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType)
    
    Dim iCurrency ,iADODBConnString
    Dim uuCurrency,uuADODBConnString

    On Error Resume Next
    
    uuADODBConnString = GetGlobalData("gADODBConnString")
    uuCurrency        = GetGlobalData("gCurrency")
    
	If Trim(uuADODBConnString) = "" Then
       iCurrency        = gCurrency
       iADODBConnString = gADODBConnString
       
       If pScreenType = "COMMONPOPUP" Then
          pScreenType  = ""
       Else
          pCookieOrNot = "COOKIE"
       End If   
       
    Else
   
       iCurrency        = uuCurrency
       iADODBConnString = uuADODBConnString
       
       If pScreenType = "COMMONPOPUP" Then
          pScreenType  = ""
       End If
	End If	
	
    If Trim(iCurrency) = "" Then
       iCurrency = Read_B_Company(iADODBConnString)
       If Trim(iCurrency) = "" Then
          iCurrency = "KRW"
       End If 
    End If
    Call QueryFormat(iCurrency, pFormType, pModuleCd,pCookieOrNot,pScreenType,iADODBConnString)    
    
End Sub    

'=======================================================================================================
'
'=======================================================================================================
Sub LoadInfTB19029B(ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType)
    
    Dim iCurrency
    
    On Error Resume Next												            '☜: Server Side Process

    iCurrency        = gCurrency

    If Trim(iCurrency) = "" Then
       iCurrency = Read_B_Company(gADODBConnString)
       If Trim(iCurrency) = "" Then
          iCurrency = "KRW"
       End If 
    End If
    Call QueryFormat(iCurrency, pFormType, pModuleCd,pCookieOrNot,pScreenType,gADODBConnString)    
    
End Sub    


'=======================================================================================================
'
'=======================================================================================================
Sub QueryFormat(ByVal pCurrency, ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType, ByVal pADODBConnString)
    Dim lgAmtOfMoneyIntegeral
    Dim lgQtyIntegeral
    Dim lgUnitCostIntegeral
    Dim lgExchRateIntegeral

    Dim lgAmtOfMoneyDecPoint
    Dim lgQtyDecPoint
    Dim lgUnitCostDecPoint
    Dim lgExchRateDecPoint
		       
    Dim lgAmtOfMoneyRndPolicy
    Dim lgQtyRndPolicy
    Dim lgUnitCostRndPolicy
    Dim lgExchRateRndPolicy
		       
    Dim lgAmtOfMoneyRndUnit
    Dim lgQtyRndUnit
    Dim lgUnitCostRndUnit
    Dim lgExchRateRndUnit

    Dim iRET
    Dim A,B,C
    Dim iTmp
    Dim iMoneyDefault, iQtyDefault, iUnitCostDefault, iExchRateDefault
    Dim iMoneyIntegeralDefault, iQtyIntegeralDefault, iUnitCostIntegeralDefault, iExchRateIntegeralDefault
    Dim iUrlTemp
    
    On Error Resume Next
  
    lgAmtOfMoneyIntegeral   = ""
    lgQtyIntegeral          = ""
    lgUnitCostIntegeral     = ""
    lgExchRateIntegeral     = ""

    lgAmtOfMoneyDecPoint    = ""
    lgQtyDecPoint           = ""
    lgUnitCostDecPoint      = ""
    lgExchRateDecPoint      = ""
		       
    lgAmtOfMoneyRndPolicy   = ""
    lgQtyRndPolicy          = ""
    lgUnitCostRndPolicy     = ""
    lgExchRateRndPolicy     = ""

    lgAmtOfMoneyRndUnit     = ""
    lgQtyRndUnit            = ""
    lgUnitCostRndUnit       = ""
    lgExchRateRndUnit       = ""
    
    iRET = ReadB_Numeric_format("2",pCurrency, pFormType, pModuleCd,iTmp,pADODBConnString)   'About Amount of Money
    
    If iRET = True Then
       iTmp = Split(iTmp,vbTab)
       A = iTmp(0)
       B = iTmp(1)
       C = iTmp(2)
    ElseIf iRET = False Then
       iRET = ReadB_Numeric_format("2",pCurrency, pFormType, "*" ,iTmp,pADODBConnString)
       If iRET = True Then
          iTmp = Split(iTmp,vbTab)
          A = iTmp(0)
          B = iTmp(1)
          C = iTmp(2)
       ElseIf iRET = False Then
          iRET = ReadB_Numeric_format("2",pCurrency, "I", "*" ,iTmp,pADODBConnString)
          If iRET = True Then
             iTmp = Split(iTmp,vbTab)
             A = iTmp(0)
             B = iTmp(1)
             C = iTmp(2)
          ElseIf iRET = False Then
             A =  "2"
             B =  "0.001" 
             C =  "3"
          End If
       End If
    End If
    
    lgAmtOfMoneyDecPoint  =  A
    lgAmtOfMoneyRndUnit   =  B
    lgAmtOfMoneyRndPolicy =  C    
    
    iRET = ReadB_Numeric_format("4",pCurrency, pFormType, pModuleCd,iTmp,pADODBConnString)   'About Unit Cost
    
    If iRET = True Then
       iTmp = Split(iTmp,vbTab)
       A = iTmp(0)
       B = iTmp(1)
       C = iTmp(2)
    ElseIf iRET = False Then
       iRET = ReadB_Numeric_format("4",pCurrency, pFormType, "*" ,iTmp,pADODBConnString)
       If iRET = True Then
          iTmp = Split(iTmp,vbTab)
          A = iTmp(0)
          B = iTmp(1)
          C = iTmp(2)
       ElseIf iRET = False Then
          iRET = ReadB_Numeric_format("4",pCurrency, "I", "*" ,iTmp,pADODBConnString)
          If iRET = True Then
             iTmp = Split(iTmp,vbTab)
             A = iTmp(0)
             B = iTmp(1)
             C = iTmp(2)
          ElseIf iRET = False Then
             A =  "4"
             B =  "0.00001" 
             C =  "3"
          End If
       End If
    End If
    
    lgUnitCostDecPoint  =  A
    lgUnitCostRndUnit   =  B
    lgUnitCostRndPolicy =  C

    iRET = ReadB_Numeric_format("5",pCurrency, pFormType, pModuleCd,iTmp,pADODBConnString)          'About Exchange Rate
    
    If iRET = True Then
       iTmp = Split(iTmp,vbTab)
       A = iTmp(0)
       B = iTmp(1)
       C = iTmp(2)
    ElseIf iRET = False Then
       iRET = ReadB_Numeric_format("5",pCurrency, pFormType, "*" ,iTmp,pADODBConnString)
       If iRET = True Then
          iTmp = Split(iTmp,vbTab)
          A = iTmp(0)
          B = iTmp(1)
          C = iTmp(2)
       ElseIf iRET = False Then
          iRET = ReadB_Numeric_format("5",pCurrency, "I", "*" ,iTmp,pADODBConnString)
          If iRET = True Then
             iTmp = Split(iTmp,vbTab)
             A = iTmp(0)
             B = iTmp(1)
             C = iTmp(2)
          ElseIf iRET = False Then
             A =  "6"
             B =  "0.0000001" 
             C =  "3"
          End If
       End If
    End If
    
    lgExchRateDecPoint  =  A
    lgExchRateRndUnit   =  B
    lgExchRateRndPolicy =  C

    iRET = ReadB_Count_format(pFormType, pModuleCd,iTmp,pADODBConnString)                             'About Quantity
    
    If iRET = True Then
       iTmp = Split(iTmp,vbTab)
       A = iTmp(0)
       B = iTmp(1)
       C = iTmp(2)
    ElseIf iRET = False Then
       iRET = ReadB_Count_format(pFormType, "*" ,iTmp,pADODBConnString)
       If iRET = True Then
          iTmp = Split(iTmp,vbTab)
          A = iTmp(0)
          B = iTmp(1)
          C = iTmp(2)
       ElseIf iRET = False Then
          iRET = ReadB_Count_format("I", "*" ,iTmp,pADODBConnString)
          If iRET = True Then
             iTmp = Split(iTmp,vbTab)
             A = iTmp(0)
             B = iTmp(1)
             C = iTmp(2)
          ElseIf iRET = False Then
             A =  "4"
             B =  "0.00001" 
             C =  "3"
          End If
       End If
    End If

    lgQtyDecPoint  = A
    lgQtyRndUnit   = B
    lgQtyRndPolicy = C
     
    lgAmtOfMoneyIntegeral    = 15 - CInt(lgAmtOfMoneyDecPoint)
    lgQtyIntegeral           = 15 - CInt(lgQtyDecPoint)
    lgUnitCostIntegeral      = 15 - CInt(lgUnitCostDecPoint)
    lgExchRateIntegeral      = 15 - CInt(lgExchRateDecPoint)
    
    If lgQtyIntegeral > 11 Then
       lgQtyIntegeral =  11
    End If

    If lgAmtOfMoneyIntegeral > 13 Then
       lgAmtOfMoneyIntegeral =  13
    End If

    If lgUnitCostIntegeral > 11 Then
       lgUnitCostIntegeral =  11
    End If
    
    If lgExchRateIntegeral > 9 Then
       lgExchRateIntegeral =  9
    End If

' 2003-02-19 Kim In Tae
'    Call ReadB_Default_Numeric_format(pADODBConnString ,pFormType, iMoneyDefault ,iQtyDefault, iUnitCostDefault, iExchRateDefault)
    iMoneyDefault = 2
    iQtyDefault = 4
    iUnitCostDefault = 4
    iExchRateDefault = 6
    iMoneyIntegeralDefault    = 15 - CInt(iMoneyDefault)
    iQtyIntegeralDefault      = 15 - CInt(iQtyDefault)
    iUnitCostIntegeralDefault = 15 - CInt(iUnitCostDefault)
    iExchRateIntegeralDefault = 15 - CInt(iExchRateDefault)

' z_ado 때문에 임시로 1더함 
'    iMoneyDefault    = iMoneyDefault + 1
'    iQtyDefault      = iQtyDefault + 1
'    iUnitCostDefault = iUnitCostDefault + 1
'    iExchRateDefault = iExchRateDefault + 1
    
'    If iQtyIntegeralDefault > 11 Then
'       iQtyIntegeralDefault =  11
'    End If

'    If iMoneyIntegeralDefault > 13 Then
'       iMoneyIntegeralDefault =  13
'    End If

'    If iUnitCostIntegeralDefault > 11 Then
'       iUnitCostIntegeralDefault =  11
'    End If
    
'    If iExchRateIntegeralDefault > 9 Then
'       iExchRateIntegeralDefault =  9
'    End If
    
    If pCookieOrNot = "COOKIE" Then

        Response.Cookies("unierp")("gAmtOfMoney")          = lgAmtOfMoneyDecPoint
        Response.Cookies("unierp")("gQty")                 = lgQtyDecPoint
        Response.Cookies("unierp")("gUnitCost")            = lgUnitCostDecPoint
        Response.Cookies("unierp")("gExchRate")            = lgExchRateDecPoint

        Response.Cookies("unierp")("gAmtOfMoneyRndPolicy") = lgAmtOfMoneyRndPolicy
        Response.Cookies("unierp")("gQtyRndPolicy")        = lgQtyRndPolicy
        Response.Cookies("unierp")("gUnitCostRndPolicy")   = lgUnitCostRndPolicy
        Response.Cookies("unierp")("gExchRateRndPolicy")   = lgExchRateRndPolicy
            
        Response.Cookies("unierp")("gAmtOfMoneyRndUnit")   = lgAmtOfMoneyRndUnit
        Response.Cookies("unierp")("gQtyRndUnit")          = lgQtyRndUnit
        Response.Cookies("unierp")("gUnitCostRndUnit")     = lgUnitCostRndUnit
        Response.Cookies("unierp")("gExchRateRndUnit")     = lgExchRateRndUnit

' 2003-02-19 Kim In Tae
'        Response.Cookies("unierp")("gAmtOfMoneyDefault")   = iMoneyDefault
'        Response.Cookies("unierp")("gQtyDefault")          = iQtyDefault
'        Response.Cookies("unierp")("gUnitCostDefault")     = iUnitCostDefault
'        Response.Cookies("unierp")("gExchRateDefault")     = iExchRateDefault
       
    End If   

    Response.Cookies("unierp")("gLang") = Request.Cookies("unierp")("gLang")

    iUrlTemp = request.servervariables("path_info")
    iUrlTemp = Split(iUrlTemp,"/")
    Response.Cookies("unierp").path = "/" & iUrlTemp(1) & "/" & iUrlTemp(2)


    A =     "0"                  & Chr(11)  '0  Reserved
    A = A & "0"                  & Chr(11)  '1  Reserved
    A = A & lgAmtOfMoneyDecPoint & Chr(11)  '2  
    A = A & lgQtyDecPoint        & Chr(11)  '3  
    A = A & lgUnitCostDecPoint   & Chr(11)  '4  
    A = A & lgExchRateDecPoint   & Chr(11)  '5  
    A = A & "0"                  & Chr(11)  '6  Reserved
    A = A & "0"                  & Chr(11)  '7  Reserved
    A = A & "0"                  & Chr(11)  '8  Reserved
    A = A & "0"                  & Chr(11)  '9  Reserved
    A = A & iMoneyDefault        & Chr(11)  'A
    A = A & iQtyDefault          & Chr(11)  'B
    A = A & iUnitCostDefault     & Chr(11)  'C
    A = A & iExchRateDefault     & Chr(11)  'D
        

    B =     "15"                      & Chr(11)  '0  Reserved
    B = B & "15"                      & Chr(11)  '1  Reserved
    B = B & lgAmtOfMoneyIntegeral     & Chr(11)  '2  
    B = B & lgQtyIntegeral            & Chr(11)  '3  
    B = B & lgUnitCostIntegeral       & Chr(11)  '4  
    B = B & lgExchRateIntegeral       & Chr(11)  '5  
    B = B & "15"                      & Chr(11)  '6  Reserved
    B = B & "15"                      & Chr(11)  '7  Reserved
    B = B & "15"                      & Chr(11)  '8  Reserved
    B = B & "15"                      & Chr(11)  '9  Reserved    
    B = B & iMoneyIntegeralDefault    & Chr(11)  'A
    B = B & iQtyIntegeralDefault      & Chr(11)  'B
    B = B & iUnitCostIntegeralDefault & Chr(11)  'C
    B = B & iExchRateIntegeralDefault & Chr(11)  'D
    
    Select Case pScreenType     
       Case  "BA" ,"MA" ,"OA" ,"QA" ,"PA" ,"RA" ,""
' 2003-02-19 Kim In Tae
'           Response.Write " ggAmtOfMoneyDefault.DecPoint  = " & iMoneyDefault  & vbCr
'           Response.Write " ggQtyDefault.DecPoint         = " & iQtyDefault         & vbCr
'           Response.Write " ggUnitCostDefault.DecPoint    = " & iUnitCostDefault    & vbCr
'           Response.Write " ggExchRateDefault.DecPoint    = " & iExchRateDefault    & vbCr

           Response.Write " ggAmtOfMoney.DecPoint  = " & lgAmtOfMoneyDecPoint  & vbCr
           Response.Write " ggQty.DecPoint         = " & lgQtyDecPoint         & vbCr
           Response.Write " ggUnitCost.DecPoint    = " & lgUnitCostDecPoint    & vbCr
           Response.Write " ggExchRate.DecPoint    = " & lgExchRateDecPoint    & vbCr

           Response.Write " ggAmtOfMoney.RndPolicy = " & lgAmtOfMoneyRndPolicy & vbCr
           Response.Write " ggQty.RndPolicy        = " & lgQtyRndPolicy        & vbCr
           Response.Write " ggUnitCost.RndPolicy   = " & lgUnitCostRndPolicy   & vbCr
           Response.Write " ggExchRate.RndPolicy   = " & lgExchRateRndPolicy   & vbCr

           Response.Write " ggAmtOfMoney.RndUnit   = """ & lgAmtOfMoneyRndUnit & """" & vbCr
           Response.Write " ggQty.RndUnit          = """ & lgQtyRndUnit        & """" & vbCr
           Response.Write " ggUnitCost.RndUnit     = """ & lgUnitCostRndUnit   & """" & vbCr
           Response.Write " ggExchRate.RndUnit     = """ & lgExchRateRndUnit   & """" & vbCr
	
           Response.Write " ggStrDeciPointPart     = """ & A & """" & vbCr
           Response.Write " ggStrIntegeralPart     = """ & B & """" & vbCr
       Case  "BB" ,"MB" ,"OB" ,"QB" ,"PB" ,"RB" 
           ggAmtOfMoney.DecPoint    = lgAmtOfMoneyDecPoint
           ggQty.DecPoint           = lgQtyDecPoint
           ggUnitCost.DecPoint      = lgUnitCostDecPoint
           ggExchRate.DecPoint      = lgExchRateDecPoint

           ggAmtOfMoney.RndPolicy   = lgAmtOfMoneyRndPolicy
           ggQty.RndPolicy          = lgQtyRndPolicy
           ggUnitCost.RndPolicy     = lgUnitCostRndPolicy
           ggExchRate.RndPolicy     = lgExchRateRndPolicy

           ggAmtOfMoney.RndUnit     = lgAmtOfMoneyRndUnit
           ggQty.RndUnit            = lgQtyRndUnit
           ggUnitCost.RndUnit       = lgUnitCostRndUnit
           ggExchRate.RndUnit       = lgExchRateRndUnit           
' 2003-02-19 Kim In Tae
'           ggAmtOfMoneyDefault.DecPoint = iMoneyDefault
'           ggQtyDefault.DecPoint        = iQtyDefault
'           ggUnitCostDefault.DecPoint   = iUnitCostDefault
'           ggExchRateDefault.DecPoint   = iExchRateDefault

    End Select

End Sub	
'=======================================================================================================
'
'=======================================================================================================
Sub LoadBNumericFormat(ByVal pFormType, ByVal pModuleCd)

    Dim iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy
    
    Call LoadBNumericFormatSub(pFormType, pModuleCd,"COOKIE"     ,""          ,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    
    If Mid(iDataType,1,2) = "NF" Then
       Call LoadBNumericFormatSub(pFormType, "*","COOKIE"     ,""          ,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    End If
    
    Call SetDefaultDec(iDataType,iDecimals)
    
    Response.Write " gBDataType       =  Split( """ & iDataType       & """ ,Chr(11)) " & vbCr
    Response.Write " gBCurrency       =  Split( """ & iCurrency       & """ ,Chr(11)) " & vbCr
    Response.Write " gBDecimals       =  Split( """ & iDecimals       & """ ,Chr(11)) " & vbCr
    Response.Write " gBRoundingUnit   =  Split( """ & iRoundingUnit   & """ ,Chr(11)) " & vbCr
    Response.Write " gBRoundingPolicy =  Split( """ & iRoundingPolicy & """ ,Chr(11)) " & vbCr   
    
End Sub    

'=======================================================================================================
'
'=======================================================================================================
Sub LoadBNumericFormatA(ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType)

    Dim iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy
    
    Call LoadBNumericFormatSub(pFormType, pModuleCd, pCookieOrNot, pScreenType,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    
    If Mid(iDataType,1,2) = "NF" Then
       Call LoadBNumericFormatSub(pFormType, "*", pCookieOrNot, pScreenType,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    End If

    Call SetDefaultDec(iDataType,iDecimals)    
    
    Response.Write " gBDataType       =  Split( """ & iDataType       & """ ,Chr(11)) " & vbCr
    Response.Write " gBCurrency       =  Split( """ & iCurrency       & """ ,Chr(11)) " & vbCr
    Response.Write " gBDecimals       =  Split( """ & iDecimals       & """ ,Chr(11)) " & vbCr
    Response.Write " gBRoundingUnit   =  Split( """ & iRoundingUnit   & """ ,Chr(11)) " & vbCr
    Response.Write " gBRoundingPolicy =  Split( """ & iRoundingPolicy & """ ,Chr(11)) " & vbCr
    
End Sub
'=======================================================================================================
'
'=======================================================================================================
Sub LoadBNumericFormatB(ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType)

    Dim iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy
    
    Call LoadBNumericFormatSub(pFormType, pModuleCd, pCookieOrNot, pScreenType,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    
    If Mid(iDataType,1,2) = "NF" Then
       Call LoadBNumericFormatSub(pFormType, "*", pCookieOrNot, pScreenType,iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy)
    End If

    gBDataType       =  Split(iDataType       ,Chr(11)) 
    gBCurrency       =  Split(iCurrency       ,Chr(11)) 
    gBDecimals       =  Split(iDecimals       ,Chr(11)) 
    gBRoundingUnit   =  Split(iRoundingUnit  ,Chr(11)) 
    gBRoundingPolicy =  Split(iRoundingPolicy,Chr(11)) 
    
End Sub
'=======================================================================================================
'
'=======================================================================================================
Sub LoadBNumericFormatSub(ByVal pFormType, ByVal pModuleCd, ByVal pCookieOrNot,ByVal pScreenType , _
                          pDataType,pCurrency,pDecimals,pRoundingUnit,pRoundingPolicy)

    Dim strSQL
    Dim pRs
    Dim iDataType,iCurrency,iDecimals,iRoundingUnit,iRoundingPolicy
    Dim           irCurrency,iADODBConnString
    Dim iUrlTemp
    Dim uuCurrency
    Dim uuADODBConnString

    On Error Resume Next
    
    uuCurrency        = GetGlobalData("gCurrency")
    uuADODBConnString = GetGlobalData("gADODBConnString")

	If Trim(uuADODBConnString) = "" Then
       irCurrency       = gCurrency
       iADODBConnString = gADODBConnString
       
       If pScreenType = "COMMONPOPUP" Then
          pScreenType  = ""
       Else
          pCookieOrNot = "COOKIE"
       End If   
       
    Else
       irCurrency       = uuCurrency
       iADODBConnString = uuADODBConnString
	End If	
    
    iDataType        = ""
    iCurrency        = ""
    iDecimals        = ""
    iRoundingUnit    = ""
    iRoundingPolicy  = ""
    
    Set pRs = Server.CreateObject("ADODB.Recordset")   

    strSQL = "SELECT DATA_TYPE,CURRENCY,DECIMALS,ROUNDING_UNIT,ROUNDING_POLICY FROM b_numeric_format"
    strSQL = strSQL & " WHERE FORM_TYPE = '" & pFormType & "'"
    strSQL = strSQL & " AND MODULE_CD = '" & pModuleCd & "'"
    If pModuleCd <> "*" Then
        strSQL = strSQL & " UNION SELECT a.DATA_TYPE,a.CURRENCY,a.DECIMALS,a.ROUNDING_UNIT,a.ROUNDING_POLICY FROM b_numeric_format a"
        strSQL = strSQL & " WHERE FORM_TYPE = '" & pFormType & "'"
        strSQL = strSQL & " AND MODULE_CD = '*'"
        strSQL = strSQL & " AND NOT EXISTS (SELECT b.FORM_TYPE FROM b_numeric_format b"
        strSQL = strSQL & " WHERE FORM_TYPE = '" & pFormType & "'"
        strSQL = strSQL & " AND MODULE_CD = '" & pModuleCd & "'"
        strSQL = strSQL & " AND a.data_type=b.data_type AND a.currency=b.currency)"
    End If
    If pFormType <> "I" Then
        strSQL = strSQL & " UNION SELECT c.DATA_TYPE,c.CURRENCY,c.DECIMALS,c.ROUNDING_UNIT,c.ROUNDING_POLICY FROM b_numeric_format c"
        strSQL = strSQL & " WHERE FORM_TYPE = 'I'"
        strSQL = strSQL & " AND MODULE_CD = '*'"
        strSQL = strSQL & " AND NOT EXISTS (SELECT d.FORM_TYPE FROM b_numeric_format d"
        strSQL = strSQL & " WHERE FORM_TYPE = '" & pFormType & "'"
        If pModuleCd <> "*" Then
        	strSQL = strSQL & " AND (MODULE_CD = '" & pModuleCd & "' OR MODULE_CD = '*')"
        Else
        	strSQL = strSQL & " AND MODULE_CD = '*'"
        End If
        strSQL = strSQL & " AND d.data_type=c.data_type AND d.currency=c.currency)"
    End If
	
    pRs.Open strSQL,iADODBConnString,0,1
    
    Do while Not (pRs.EOF Or pRs.BOF)
       iDataType       = iDataType       & pRs(0) & Chr(11)
       iCurrency       = iCurrency       & pRs(1) & Chr(11)
       iDecimals       = iDecimals       & pRs(2) & Chr(11)
       iRoundingUnit   = iRoundingUnit   & pRs(3) & Chr(11)
       iRoundingPolicy = iRoundingPolicy & pRs(4) & Chr(11)
       pRs.MoveNext
	Loop
	
    pRs.Close
	Set pRs = Nothing
	
    Set pRs = Server.CreateObject("ADODB.Recordset")   

    strSQL = "Select DECIMALS,ROUNDING_UNIT,ROUNDING_POLICY FROM B_COUNT_FORMAT "
    strSQL = strSQL & " WHERE MODULE_CD = '" & pModuleCd & "'"
    strSQL = strSQL & " AND   FORM_TYPE = '" & pFormType & "'"
    
    pRs.Open strSQL,iADODBConnString,0,1
    
    Do while Not (pRs.EOF Or pRs.BOF)
       iDataType       = iDataType       & "3"        & Chr(11)
       iCurrency       = iCurrency       & irCurrency & Chr(11)
       iDecimals       = iDecimals       & pRs(0)     & Chr(11)
       iRoundingUnit   = iRoundingUnit   & pRs(1)     & Chr(11)
       iRoundingPolicy = iRoundingPolicy & pRs(2)     & Chr(11)
       pRs.MoveNext
	Loop
	
    pRs.Close
	Set pRs = Nothing		

	If Trim(iDataType) = "" Then
	   iDataType = "NF" & Chr(11)
	End If
	If Trim(iCurrency) = "" Then
	   iCurrency = "NF" & Chr(11)
	End If
	If Trim(iDecimals) = "" Then
	   iDecimals = "NF" & Chr(11)
	End If
	If Trim(iRoundingUnit) = "" Then
	   iRoundingUnit = "NF" & Chr(11)
	End If
	If Trim(iRoundingPolicy) = "" Then
	   iRoundingPolicy = "NF" & Chr(11)
	End If

    If pCookieOrNot = "COOKIE" Then
       Response.Cookies("unierp")("gBDataType")        = iDataType           '0
       Response.Cookies("unierp")("gBCurrency")        = iCurrency           '1
       Response.Cookies("unierp")("gBDecimals")        = iDecimals           '2
       Response.Cookies("unierp")("gBRoundingUnit")    = iRoundingUnit       '3
       Response.Cookies("unierp")("gBRoundingPolicy")  = iRoundingPolicy     '4
    End If   

    iUrlTemp = request.servervariables("path_info")
    iUrlTemp = Split(iUrlTemp,"/")
    Response.Cookies("unierp").path = "/" & iUrlTemp(1) & "/" & iUrlTemp(2)

    pDataType        = iDataType           '0
    pCurrency        = iCurrency           '1
    pDecimals        = iDecimals           '2
    pRoundingUnit    = iRoundingUnit       '3
    pRoundingPolicy  = iRoundingPolicy     '4
    
End Sub
'===================================================================================
'
'===================================================================================
Function ReadB_Numeric_format(ByVal pDataType,ByVal pCurrency,ByVal  pFormType,ByVal  pModuleCd, pRetRec, ByVal pADODBConnString)
    Dim strSQL
    Dim adoRec
    Dim iTmp

    On Error Resume Next
    
    ReadB_Numeric_format = False

    strSQL = "          Select a.decimals,a.rounding_unit,a.rounding_policy "
    strSQL = strSQL & " From  b_numeric_format A , b_currency B "
    strSQL = strSQL & " Where A.currency   = B.currency "
    strSQL = strSQL & "   and A.data_type  = '" & pDataType & "'"
    strSQL = strSQL & "   and A.currency   = '" & pCurrency & "'"
    strSQL = strSQL & "   and A.form_type  = '" & pFormType & "'"
    strSQL = strSQL & "   and A.module_cd  = '" & pModuleCd & "'"

    Set adoRec = Server.CreateObject("ADODB.RecordSet")
    adoRec.Open strSQL, pADODBConnString,0,1
    
    If Err.number = 0 Then
       If adoRec.BOF Or adoRec.EOF Then
       Else
          pRetRec =  adoRec.GetString()
          ReadB_Numeric_format = True
       End if   
    End If
    
    adoRec.Close
    Set adoRec = Nothing

End Function

'===================================================================================
'
'===================================================================================
Function Read_B_Company(ByVal pADODBConnString)
    Dim strSQL
    Dim adoRec
    Dim iTmp
    
    On Error Resume Next
    
    Read_B_Company = ""

    Set adoRec = Server.CreateObject("ADODB.RecordSet")

    strSQL = " Select loc_cur From b_company"
    
    adoRec.Open strSQL, pADODBConnString,0,1
    
    If Err.number = 0 Then
       If adoRec.BOF Or adoRec.EOF Then
       Else
          If Not Isnull(adoRec(0)) Then
             iTmp = adoRec(0)
          End If
       End If
    End If

    If Trim(iTmp) = "" Then
       iTmp = "KRW"
    End If 
    Read_B_Company = iTmp
      
    adoRec.Close
    Set adoRec = Nothing
    
End Function

'===================================================================================
'
'===================================================================================
Function ReadB_Count_format(ByVal pFormType, ByVal pModuleCd,pRetRec, ByVal pADODBConnString)
    Dim strSQL
    Dim adoRec
    
    On Error Resume Next
    
    ReadB_Count_format = False

    strSQL = "          Select decimals,rounding_unit,rounding_policy From  b_count_format"
    strSQL = strSQL & " Where module_cd  = '" & pModuleCd & "' and form_type  = '" & pFormType & "'"

    Set adoRec = Server.CreateObject("ADODB.RecordSet")
    adoRec.Open strSQL, pADODBConnString,0,1

    If Err.number = 0 Then
       If adoRec.BOF Or adoRec.EOF Then
       Else
          pRetRec =  adoRec.GetString()
          ReadB_Count_format = True
       End if   
    End If
    
    adoRec.Close
    Set adoRec = Nothing

End Function

'===================================================================================
'
'===================================================================================
Function ReadB_Default_Numeric_format(ByVal pADODBConnString ,ByVal pFormType, pAmtOfMoney ,pQty, pUnitCost, pExchRate  )
    Dim strSQL
    Dim adoRec

    On Error Resume Next
    
    ReadB_Default_Numeric_format = False

    pAmtOfMoney = 2
    pQty        = 4
    pUnitCost   = 4
    pExchRate   = 6

    strSQL = "          Select data_type , max(decimals) "
    strSQL = strSQL & " From  b_numeric_format Where FORM_TYPE='" & pFormType & "'"
    strSQL = strSQL & " Group By data_type "

    Set adoRec = Server.CreateObject("ADODB.RecordSet")
    adoRec.Open strSQL, pADODBConnString,0,1
    
    If Err.number = 0 Then
        If adoRec.BOF Or adoRec.EOF Then
        Else
            Do while Not (adoRec.EOF Or adoRec.BOF)
                Select Case adoRec(0)
                    Case "2"
                        pAmtOfMoney = adoRec(1)
                    Case "4"
                        pUnitCost = adoRec(1)
                    Case "5"
                        pExchRate = adoRec(1)
                End Select
                adoRec.MoveNext
            Loop
            pQty = 0
            ReadB_Default_Numeric_format = True
       End if   
    End If
    
    adoRec.Close
    Set adoRec = Nothing

End Function
Sub SetDefaultDec(Byval pDataType,Byval pDecimals)
    Dim iMaxdecimal(6)
    Dim i, pArrDataType, pArrDecimals, iDTTemp, iDecTemp
    
    iMaxdecimal(2) = -1
    iMaxdecimal(3) = -1
    iMaxdecimal(4) = -1
    iMaxdecimal(5) = -1

    pArrDataType = split(pDataType,chr(11))
    pArrDecimals = split(pDecimals,chr(11))
    
    For i = 0 to UBound(pArrDataType) - 1
        iDTTemp = CInt(pArrDataType(i))
        iDecTemp = CInt(pArrDecimals(i))
	    If iMaxdecimal(iDTTemp) < iDecTemp Then
	        iMaxdecimal(iDTTemp) = iDecTemp
	    End If
    Next

    If iMaxdecimal(2) = -1 Or iMaxdecimal(2) > 2 Then
        iMaxdecimal(2) = 2
    End If    
    If iMaxdecimal(3) = -1 Or iMaxdecimal(3) > 4 Then
        iMaxdecimal(3) = 4
    End If    
    If iMaxdecimal(4) = -1 Or iMaxdecimal(4) > 4 Then
        iMaxdecimal(4) = 4
    End If    
    If iMaxdecimal(5) = -1 Or iMaxdecimal(5) > 6 Then
        iMaxdecimal(5) = 6
    End If    

    Response.Write " Dim iArrTemp " & vbCr
    Response.Write " iArrTemp = split(ggStrDeciPointPart,Chr(11)) " & vbCr
     
    Response.Write " iArrTemp(10) = """ & iMaxdecimal(2) & """" & vbCr
    Response.Write " iArrTemp(11) = """ & iMaxdecimal(3) & """" & vbCr
    Response.Write " iArrTemp(12) = """ & iMaxdecimal(4) & """" & vbCr
    Response.Write " iArrTemp(13) = """ & iMaxdecimal(5) & """" & vbCr
    Response.Write " ggStrDeciPointPart = Join(iArrTemp , chr(11)) " & vbCr
End Sub


Function GetGlobalData(pData)   '2003-08-07 leejinsoo

    Dim EDCodeComEDCodeObj1
    Dim xmlDOMDocument
    
    On Error Resume Next

    Set xmlDOMDocument = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDOMDocument.async = False 
	    
	xmlDOMDocument.Load (Trim(Request.Cookies("unierp")("gXMLFileNm")))

    Set EDCodeComEDCodeObj1 = Server.CreateObject("EDCodeCom.EDCodeObj.1")
    
    xmlDOMDocument.LoadXML (Replace(EDCodeComEDCodeObj1.Decode(xmlDOMDocument.firstChild.firstChild.xml),vbCrLf,""))    		

    GetGlobalData = xmlDOMDocument.selectSingleNode("/uniERP/LoadBasisGlobalInf/" & pData ).text

    Set xmlDOMDocument      = Nothing
    Set EDCodeComEDCodeObj1 = Nothing

End Function

%>

