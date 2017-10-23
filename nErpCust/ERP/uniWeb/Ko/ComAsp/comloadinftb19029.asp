<%							
'**********************************************************************************************
'*  1. Module Name          : ComLoadInfTB19029
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 1999/09/10
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              : Do not touch this file except admin
'*                          : Called by reference or commonpopup
'**********************************************************************************************

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
		       
		       Dim lgDataType
               Dim lgCurrency
               Dim lgDecimals
               Dim lgRoundingUnit
               Dim lgRoundingPolicy
    
		       lgAmtOfMoneyDecPoint  = Trim(Request.Cookies("unierp")("gAmtOfMoney"))
		       lgQtyDecPoint         = Trim(Request.Cookies("unierp")("gQty"))
		       lgUnitCostDecPoint    = Trim(Request.Cookies("unierp")("gUnitCost"))
		       lgExchRateDecPoint    = Trim(Request.Cookies("unierp")("gExchRate"))
		       
		       If lgAmtOfMoneyDecPoint  = "" Then
		          lgAmtOfMoneyDecPoint  = 2
		       End If
		       If lgQtyDecPoint         = "" Then
		          lgQtyDecPoint         = 4
		       End If
		       If lgUnitCostDecPoint    = "" Then
		          lgUnitCostDecPoint    = 4
		       End If
		       If lgExchRateDecPoint    = "" Then
		          lgExchRateDecPoint    = 6
		       End If
		       
              lgAmtOfMoneyRndPolicy  = Trim(Request.Cookies("unierp")("gAmtOfMoneyRndPolicy"))
              lgQtyRndPolicy         = Trim(Request.Cookies("unierp")("gQtyRndPolicy"))
              lgUnitCostRndPolicy    = Trim(Request.Cookies("unierp")("gUnitCostRndPolicy"))
              lgExchRateRndPolicy    = Trim(Request.Cookies("unierp")("gExchRateRndPolicy"))
              
              If lgAmtOfMoneyRndPolicy  = "" Then
                 lgAmtOfMoneyRndPolicy  = 3
              End If
              If lgQtyRndPolicy         = "" Then
                 lgQtyRndPolicy         = 3
              End If
              If lgUnitCostRndPolicy    = "" Then
                 lgUnitCostRndPolicy    = 3
              End If
              If lgExchRateRndPolicy    = "" Then
                 lgExchRateRndPolicy    = 3              
              End If
              
    
              lgAmtOfMoneyRndUnit    = Trim(Request.Cookies("unierp")("gAmtOfMoneyRndUnit"))
              lgQtyRndUnit           = Trim(Request.Cookies("unierp")("gQtyRndUnit"))
              lgUnitCostRndUnit      = Trim(Request.Cookies("unierp")("gUnitCostRndUnit"))
              lgExchRateRndUnit      = Trim(Request.Cookies("unierp")("gExchRateRndUnit"))
              
              If lgAmtOfMoneyRndUnit  = "" Then
                 lgAmtOfMoneyRndUnit  = 0.001
              End If
              If lgQtyRndUnit         = "" Then
                 lgQtyRndUnit         = 0.00001
              End If
              If lgUnitCostRndUnit    = "" Then
                 lgUnitCostRndUnit    = 0.00001
              End If
              If lgExchRateRndUnit    = "" Then
                 lgExchRateRndUnit    = 0.0000001
              End If		       
              
              
		      lgAmtOfMoneyIntegeral = 15 - CInt(lgAmtOfMoneyDecPoint) 
		      lgQtyIntegeral        = 15 - CInt(lgQtyDecPoint) 
		      lgUnitCostIntegeral   = 15 - CInt(lgUnitCostDecPoint) 
		      lgExchRateIntegeral   = 15 - CInt(lgExchRateDecPoint) 
              
		       
		    %>

            ggAmtOfMoney.DecPoint  = <%=lgAmtOfMoneyDecPoint%>                            '¢Ð: Decimal point place
            ggQty.DecPoint         = <%=lgQtyDecPoint%>                           '¢Ð:
            ggUnitCost.DecPoint    = <%=lgUnitCostDecPoint%>                              '¢Ð:
            ggExchRate.DecPoint    = <%=lgExchRateDecPoint%>                              '¢Ð:

            ggAmtOfMoney.RndPolicy = <%=lgAmtOfMoneyRndPolicy%>
            ggQty.RndPolicy        = <%=lgQtyRndPolicy%>
            ggUnitCost.RndPolicy   = <%=lgUnitCostRndPolicy%>
            ggExchRate.RndPolicy   = <%=lgExchRateRndPolicy%>
    
            ggAmtOfMoney.RndUnit   = <%=lgAmtOfMoneyRndUnit%>
            ggQty.RndUnit          = <%=lgQtyRndUnit%>
            ggUnitCost.RndUnit     = <%=lgUnitCostRndUnit%>
            ggExchRate.RndUnit     = <%=lgExchRateRndUnit%>

	
            ggStrDeciPointPart    =                      ""                    & gColSep  '0  Reserved
            ggStrDeciPointPart    = ggStrDeciPointPart & ""                    & gColSep  '1  Reserved
            ggStrDeciPointPart    = ggStrDeciPointPart & ggAmtOfMoney.DecPoint & gColSep  '2  
            ggStrDeciPointPart    = ggStrDeciPointPart & ggQty.DecPoint        & gColSep  '3  
            ggStrDeciPointPart    = ggStrDeciPointPart & ggUnitCost.DecPoint   & gColSep  '4  
            ggStrDeciPointPart    = ggStrDeciPointPart & ggExchRate.DecPoint   & gColSep  '5  
            ggStrDeciPointPart    = ggStrDeciPointPart & ""                    & gColSep  '7  Reserved
            ggStrDeciPointPart    = ggStrDeciPointPart & ""                    & gColSep  '8  Reserved
            ggStrDeciPointPart    = ggStrDeciPointPart & ""                    & gColSep  '9  Reserved
            ggStrDeciPointPart    = ggStrDeciPointPart & "2"                   & gColSep  'A
            ggStrDeciPointPart    = ggStrDeciPointPart & "4"                   & gColSep  'B
            ggStrDeciPointPart    = ggStrDeciPointPart & "4"                   & gColSep  'C
            ggStrDeciPointPart    = ggStrDeciPointPart & "6"                   & gColSep  'D
            
            ggStrIntegeralPart    = ""                 & ""                           & gColSep  '0  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & ""                           & gColSep  '1  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & "<%=lgAmtOfMoneyIntegeral%>" & gColSep  '2  
            ggStrIntegeralPart    = ggStrIntegeralPart & "<%=lgQtyIntegeral%>"        & gColSep  '3  
            ggStrIntegeralPart    = ggStrIntegeralPart & "<%=lgUnitCostIntegeral%>"   & gColSep  '4  
            ggStrIntegeralPart    = ggStrIntegeralPart & "<%=lgExchRateIntegeral%>"   & gColSep  '5  
            ggStrIntegeralPart    = ggStrIntegeralPart & ""                           & gColSep  '6  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & ""                           & gColSep  '7  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & ""                           & gColSep  '8  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & ""                           & gColSep  '9  Reserved
            ggStrIntegeralPart    = ggStrIntegeralPart & "13"                         & gColSep  'A  
            ggStrIntegeralPart    = ggStrIntegeralPart & "11"                         & gColSep  'B  
            ggStrIntegeralPart    = ggStrIntegeralPart & "11"                         & gColSep  'C  
            ggStrIntegeralPart    = ggStrIntegeralPart & "9"                          & gColSep  'D  

            gComNum1000 = "<%=gComNum1000%>"
            gComNumDec  = "<%=gComNumDec%>"

            
            If Trim(gComNumDec) ="" Then
               gComNumDec = "."
               gComNum1000 = ","
            End If

            <%
            lgDataType       = Trim(Request.Cookies("unierp")("gBDataType"))
            lgCurrency       = Trim(Request.Cookies("unierp")("gBCurrency"))
            lgDecimals       = Trim(Request.Cookies("unierp")("gBDecimals"))
            lgRoundingUnit   = Trim(Request.Cookies("unierp")("gBRoundingUnit"))
            lgRoundingPolicy = Trim(Request.Cookies("unierp")("gBRoundingPolicy"))
            
            If Trim(lgDataType) = "" Then
	           lgDataType = "NF" & Chr(11)
            End If
            
            If Trim(lgCurrency) = "" Then
               lgCurrency = "NF" & Chr(11)
            End If
            
            If Trim(lgDecimals) = "" Then
               lgDecimals = "NF" & Chr(11)
            End If
            
            If Trim(lgRoundingUnit) = "" Then
               lgRoundingUnit = "NF" & Chr(11)
            End If
            
            If Trim(lgRoundingPolicy) = "" Then
               lgRoundingPolicy = "NF" & Chr(11)
            End If
            %>

            gBDataType       = "<%=lgDataType%>"                    '0
            gBCurrency       = "<%=lgCurrency%>"                    '1
            gBDecimals       = "<%=lgDecimals%>"                    '2
            gBRoundingUnit   = "<%=lgRoundingUnit%>"                '3
            gBRoundingPolicy = "<%=lgRoundingPolicy%>"              '4

            gBDataType       =  Split(gBDataType      ,Chr(11)) 
            gBCurrency       =  Split(gBCurrency      ,Chr(11)) 
            gBDecimals       =  Split(gBDecimals      ,Chr(11)) 
            gBRoundingUnit   =  Split(gBRoundingUnit  ,Chr(11)) 
            gBRoundingPolicy =  Split(gBRoundingPolicy,Chr(11)) 

