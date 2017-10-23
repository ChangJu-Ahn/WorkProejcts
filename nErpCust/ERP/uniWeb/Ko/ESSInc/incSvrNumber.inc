<%
'=== gAPNum1000
'=== gAPNumDec
'=== gBConfMinorCD
'=== gBDecimals
'=== gBRoundingUnit
'=== gBCurrency
'=== gBDataType
'=== gBRoundingPolicy
'=== gComNum1000
'=== gComNumDec
'=== gCurrency
'=== gLocRndPolicy
'=== gTaxRndPolicy

'==============================================================================
'  Numeric Format Information  Reference Constants in Master Data
'==============================================================================


Const ggAmtOfMoneyNo    = "2"      ' Amount No
Const ggQtyNo           = "3"      ' Quantity No
Const ggUnitCostNo      = "4"      ' Cost No
Const ggExchRateNo      = "5"      ' Exchange Rate No

'==============================================================================
' Conversion rule no for Rnd Policy
'==============================================================================
Const gTaxRndPolicyNo   = "1" 
Const gLocRndPolicyNo   = "2" 

'========================================================================================
'Global variable for numeric format  / added on 2001/11/28 
'========================================================================================
Dim gBDataType
Dim gBCurrency
Dim gBDecimals
Dim gBRoundingUnit
Dim gBRoundingPolicy

'========================================================================================
' Decial place for Money,UnitCost,Amount,Exchangerate
'========================================================================================
Class TB19029
     Dim DecPoint			            'Decimal place point
     Dim RndPolicy                      'Round Policy
     Dim RndUnit                        'Round Unit
End Class

'Dim ggAmtExOfMoney                        'Amount of Money (For Hermes)
'Set ggAmtExOfMoney = New TB19029


Dim ggAmtOfMoney                        'Amount of Money
Set ggAmtOfMoney = New TB19029

Dim ggQty                               'Amount
Set ggQty        = New TB19029

Dim ggUnitCost                          'UnitCost
Set ggUnitCost   = New TB19029

Dim ggExchRate                          'ExchangeRate
Set ggExchRate   = New TB19029

'==============================================================================
' Name : UNICInt
' Desc : AP(DB)[String] -> AP(DB)[Integer]
'==============================================================================
Function UNICInt(ByVal pNum, ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault     
    End If
    
    UNICInt = CInt(Trim(pNum))
	
End Function

'==============================================================================
' Name : UNICDbl
' Desc : AP(DB)[String] -> AP(DB)[Double]
'==============================================================================
Function UNICDbl(ByVal pNum, ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault     
    End If

    UNICDbl = CDbl(Trim(pNum))
	
End Function

'==============================================================================
' Name : UNICLng
' Desc : AP(DB)[String] -> AP(DB)[Long]
'==============================================================================
Function UNICLng(ByVal pNum, ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault     
    End If

    UNICLng = CLng(Trim(pNum))

End Function

'==============================================================================
' Name  : UNIConvNum
' Desc  : Company -> DB
'==============================================================================
Function UNIConvNum(ByVal pNum, ByVal pDefault)
    
    If ValidateData(pNum,"SEN") = False Then
       UNIConvNum = pDefault
       Exit Function
    End If
   
    pNum = Replace(pNum , " "        , "")
    pNum = Replace(pNum , gComNum1000, "")
    
    UNIConvNum = Replace(pNum, gComNumDec , ".")

End Function

'==============================================================================
' Name  : uniConvNumStringToDouble
' Desc  : DB(String) -> DB(Double)
'==============================================================================
Function uniConvNumStringToDouble(ByVal pNum, ByVal pDefault)

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault
    End If
   
    uniConvNumStringToDouble = CDbl(Trim(pNum))
	
End Function

'==============================================================================
' Name  : UNINumClientFormat
' Desc  : DB -> Company 
'         2002/08/22 - 1000 단위 보여주기 추가 
'==============================================================================
Function UNINumClientFormat(ByVal pNum, ByVal pDecPoint, ByVal pDefault)
        
    If ValidateData(pNum,"SEN") = False Then
       UNINumClientFormat = pDefault
       Exit Function
    End If
   
    If IsNumeric(pDecPoint) = True Then
       pDecPoint = CInt(pDecPoint)    
    Else
       pDecPoint = 0
    End If
    
    If  pDecPoint  > 0 Then
        pNum = FormatNumber(pNum, pDecPoint)  
    End If
    
    pNum = MakeExpNumToStrNum(pNum)

    UNINumClientFormat = uniConvNumAToB(pNum,gAPNum1000,gAPNumDec,gComNum1000,gComNumDec,True,"X","X")
        
End Function

'=================================== CurNumClientFormat() ======================================
' Function Name : UNINumClientFormatByCurrency
' Function Desc : This method initializes Numeric Format property as accounting to Currency
' Modified      : 2002/05/10 by kim hyngsuk with lee jinsoo
'                 2002/08/23 by lee jin soo
'===============================================================================================
Function UNINumClientFormatByCurrency(ByVal pNum,ByVal pCurrency,ByVal pDataType)

    UNINumClientFormatByCurrency = UNIConvNumDBToCompanyByCurrency(pNum ,pCurrency ,pDataType ,"X" ,"X")

End Function

'===============================================================================================
'   Event Name : UNINumClientFormatByTax
'   Event Desc : This function initializes Numeric Format, which converts to Tax Format according to 
'                tax rounding policy and amount decimal policy 
' Modified     : 2002/05/10 by kim hyngsuk with lee jinsoo
'                 2002/08/23 by lee jin soo
'===============================================================================================
Function UNINumClientFormatByTax(ByVal pNum,ByVal pCurrency,ByVal pDataType)

    If ValidateData(pDataType,"SEN") = False Then
	   pDataType = ggAmtOfMoneyNo
	End if
	
    UNINumClientFormatByTax = UNIConvNumDBToCompanyByCurrency(pNum ,pCurrency ,pDataType ,gTaxRndPolicyNo ,"X")
	
End function

'===============================================================================================
' Function Name : UNIConvNumDBToCompanyByCurrency
' Function Desc :
' Modified      : 2002/05/10 by kim hyngsuk with lee jinsoo
'===============================================================================================
Function UNIConvNumDBToCompanyByCurrency(ByVal pNum,ByVal pCurrency,ByVal pDataType,ByVal pOpt1, ByVal pOpt2)

	Dim iDx
	
    UNIConvNumDBToCompanyByCurrency = ""
    
	iDx = FindIndexOfCurrency(pCurrency,pDataType)

	If CInt(iDx) < 0 Then 
       iDx = FindIndexOfCurrency(gCurrency,pDataType)
	   If CInt(iDx) < 0 Then 
	      UNIConvNumDBToCompanyByCurrency = UniConvNumDBToCompanyWithOutChange(pNum,"")   ' 2002/08/13 lee jin soo
	      Exit Function
	   End If   
	End If
	
    Select Case pOpt1
    
        Case gTaxRndPolicyNo :  UNIConvNumDBToCompanyByCurrency = UniConvNumberDBToCompany(pNum,gBDecimals(iDx),gTaxRndPolicy        ,gBRoundingUnit(iDx),0) ' 세금계산 정책 

        Case gLocRndPolicyNo :  '외화 환산 정책 

                             If gBConfMinorCD = "1" Then
                                UNIConvNumDBToCompanyByCurrency = UniConvNumberDBToCompany(pNum,gBDecimals(iDx),gBRoundingPolicy(iDx),gBRoundingUnit(iDx),0)
                             Else    
                                UNIConvNumDBToCompanyByCurrency = UniConvNumberDBToCompany(pNum,gBDecimals(iDx),gLocRndPolicy        ,gBRoundingUnit(iDx),0)
                             End If

        Case Else            :  UNIConvNumDBToCompanyByCurrency = UniConvNumberDBToCompany(pNum,gBDecimals(iDx),gBRoundingPolicy(iDx),gBRoundingUnit(iDx),0)

    End Select      

End Function

'===============================================================================================
' Function Name : FindIndexOfCurrency
' Function Desc : This method initializes Type property as accounting to Currency
'===============================================================================================
Function FindIndexOfCurrency(ByVal pCurrency, ByVal pDataType)

	Dim iDx
	
    FindIndexOfCurrency = -1

	pCurrency = UCase(Trim(pCurrency))
	
	If pDataType = "3" Then
	   pCurrency = gCurrency
	End If   

	For iDx = 0 to UBound(gBCurrency)
		If UCASE(Trim(gBCurrency(iDx))) = pCurrency Then
           If gBDataType(iDx) = pDataType Then
              FindIndexOfCurrency = iDx
              Exit Function
           End If
        End If
	Next

End Function

'==============================================================================
' Name  : UniConvNumberDBToCompany
' Desc  : DB -> Company according to the round policy
'==============================================================================
Function UniConvNumberDBToCompany(ByVal pNum,ByVal pDecPoint,ByVal pRndPolicy,ByVal pRndUnit,ByVal pDefault)

    Dim iTmpNum
    Dim iTmpNumArr
    Dim iTmpNumInt,iTmpNumDec
    
    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault
    End If
	
	If CDbl(pNum) = 0 Then
       UniConvNumberDBToCompany = Replace(FormatNumber(0, pDecPoint, -2), gAPNumDec , gComNumDec)
	   Exit Function
	End If
	
    pNum = MakeExpNumToStrNum(pNum)	

    pNum = Replace(pNum    , " "        , "")
    pNum = Replace(pNum    , gAPNum1000 , "")

    pNum= FncRoundData(pNum,pDecPoint,pRndPolicy,pRndUnit,gAPNumDec)

    UniConvNumberDBToCompany = uniConvNumAToB(pNum,gAPNum1000,gAPNumDec,gComNum1000,gComNumDec,True,"X","X")

End Function


'==============================================================================
' Name  : UniConvNumDBToCompanyWithOutChange
' Desc  : DB -> Company without round policy
'==============================================================================
Function UniConvNumDBToCompanyWithOutChange(ByVal pNum, ByVal pDefault)

    On Error Resume Next

    If ValidateData(pNum,"SEN") = False Then
       pNum = pDefault
    End If

    pNum = MakeExpNumToStrNum(pNum)	
	
    UniConvNumDBToCompanyWithOutChange = uniConvNumAToB(pNum,gAPNum1000,gAPNumDec,gComNum1000,gComNumDec,True,"X","X")
    
End Function

'========================================================================================
' Function Name : FncRoundData
' Function Desc : Round data
'========================================================================================
Function FncRoundData(ByVal pNum, ByVal pDecPoint, ByVal pRndPolicy, ByVal pRndUnit, ByVal pNumDec)
   Dim iNumARR(30)
   Dim iNumSTR
   Dim ii, jj
   Dim iBit
   Dim iNumDFindPlace
   Dim iCountAfterDec
   Dim iMeetDPoint
   Dim iPositive
   Dim iDataMaxLength
   
   If CDbl(pNum) > 0 Then
      iPositive = True
   Else
      iPositive = False
   End If

   If IsNumeric(pDecPoint) Then
      pDecPoint = CInt(pDecPoint)
   Else
      pDecPoint = 0
   End If
    
   If InStr(CStr(pNum), pNumDec) = 0 Then           'If it is Integer
      If pDecPoint > 0 Then
         FncRoundData = pNum & pNumDec & Mid("000000", 1, pDecPoint)
      Else
         FncRoundData = pNum
      End If
      Exit Function
   End If

   jj = 0
   iNumARR(jj) = 0
   iCountAfterDec = 0
   iMeetDPoint = False
   iNumDFindPlace = 0
   
   For ii = 1 To Len(pNum)
       If iMeetDPoint = True Then
          iCountAfterDec = iCountAfterDec + 1
          If iCountAfterDec > pDecPoint + 1 Then
              Exit For
          End If
       End If
       iBit = Mid(pNum, ii, 1)
       If iBit = pNumDec Then
          iMeetDPoint = True
          iNumDFindPlace = jj + 1

       ElseIf IsNumeric(iBit) = True Then
          jj = jj + 1
          iNumARR(jj) = iBit
       End If
   Next
   
   If iCountAfterDec < pDecPoint Then
      FncRoundData = pNum & Mid("000000", 1, pDecPoint - iCountAfterDec)
      Exit Function
   End If
   
   iDataMaxLength = iNumDFindPlace + pDecPoint
   Select Case pRndPolicy
         Case "1":                                                  ' 올림 
                  If CInt(iNumARR(iDataMaxLength)) > 0 Then
                     iNumARR(iDataMaxLength - 1) = iNumARR(iDataMaxLength - 1) + 1
                  End If
         Case "2"                                                   ' 내림 
         Case "3"                                                   ' 반올림 
                  If CInt(iNumARR(iNumDFindPlace + pDecPoint)) > 4 Then
                     iNumARR(iDataMaxLength - 1) = iNumARR(iDataMaxLength - 1) + 1
                  End If
   End Select
   
   iNumARR(iNumDFindPlace + pDecPoint) = 0
   
   For ii = iDataMaxLength To 1 Step -1
       If iNumARR(ii) > 9 Then
          iNumARR(ii) = iNumARR(ii) - 10
          iNumARR(ii - 1) = iNumARR(ii - 1) + 1
       End If
   Next
   
   iNumSTR = ""
   
   If iNumARR(0) > 0 Then
       iNumSTR = iNumARR(0) & iNumSTR
   End If
   
   For ii = 1 To iDataMaxLength - 1
       If iNumDFindPlace = ii Then
          iNumSTR = iNumSTR & pNumDec
       End If
       iNumSTR = iNumSTR & iNumARR(ii)
   Next
   
   If iPositive = False Then
      FncRoundData = "-" & iNumSTR
   Else
      FncRoundData = iNumSTR
   End If
End Function

'==============================================================================
' Name : MarkNum1000SEP
' Desc : MarkNum1000SEP
' Born : 2002/08/22
'==============================================================================
Function MarkNum1000SEP(ByVal pNum,ByVal pALTNum,ByVal p1000SEP)  

    Dim iData
    
    MarkNum1000SEP = ""
    
    If ValidateData(pNum,"SENF") = False Then
       MarkNum1000SEP = pALTNum       
       Exit Function
    End If
    
    If Len(Trim(pNum)) < 4 Then
       MarkNum1000SEP = pNum  
       Exit Function
    End If    

    pNum = StrReverse(pNum)
    
    Do While 1
  
       iData = iData & Mid(pNum,1,3)
       pNum  = Mid(pNum,4)

       If Trim(pNum) = "" Then
          Exit Do              ' Exit loop.
       End If
       iData       = iData & p1000SEP 
    Loop
    
    iData          = StrReverse(iData)

    If Mid(iData,1,2) = "-," Then
       MarkNum1000SEP = "-" & Mid(iData,3)
       Exit Function
    End If
    
    If Mid(iData,1,2) = "-." Then
       MarkNum1000SEP = "-" & Mid(iData,3)
       Exit Function
    End If
    
    MarkNum1000SEP = iData
   
End Function   

'==============================================================================
' Name : ValidateData
' Desc : Validate Data
' Born : 2002/08/22
'==============================================================================
Function ValidateData(ByVal pNum ,ByVal pOpt )

    ValidateData = False 
    
    If InStr(pOpt,"E") > 0 Then     
       If IsEmpty(pNum) Then
          Exit Function
       End If
    End If   

    If InStr(pOpt,"N") > 0 Then     
       If IsNull(pNum) Then
          Exit Function
       End If
    End If   
    
    If InStr(pOpt,"S") > 0 Then     
       If Trim(pNum) = "" Then
          Exit Function
       End If
    End If   
    
    If InStr(pOpt,"F") > 0 Then
       If Not IsNumeric(pNum) Then
          Exit Function
       End If    
    End If    
    
    ValidateData = True
      
End Function 

'==============================================================================
' Name : uniConvNumAToB
' Desc : Conv A to B
' Born : 2002/08/22
'==============================================================================
Function uniConvNumAToB(ByVal pNum, ByVal pNum1000From, ByVal pNumDecFrom, ByVal pNum1000To, ByVal pNumDecTo, ByVal p1000SEP,ByVal pOpt1,ByVal pOpt2)
    Dim iTmpNumInt
    Dim iTmpNumDec
    Dim iTmpNumArr

    pNum = Replace(pNum, " "         , "")
    pNum = Replace(pNum, pNum1000From, "")

    If InStr(pNum, pNumDecFrom) > 0 Then
       iTmpNumArr = Split(pNum, pNumDecFrom)
       iTmpNumInt = iTmpNumArr(0)
       iTmpNumDec = iTmpNumArr(1)
    Else
       iTmpNumInt = pNum
       iTmpNumDec = ""
    End If
    
    If p1000SEP = True Then
       If InStr(iTmpNumInt, pNum1000From) > 0 Then
          iTmpNumInt = Replace(iTmpNumInt, pNum1000From, pNum1000To)
       Else
          iTmpNumInt = MarkNum1000SEP(iTmpNumInt, "0", pNum1000To)
       End If
    End If
    
    If iTmpNumDec = "" Then
       uniConvNumAToB = iTmpNumInt
    Else
       uniConvNumAToB = iTmpNumInt & pNumDecTo & iTmpNumDec
    End If

End Function

'=========================================================================================================================
'
'
'=========================================================================================================================
Function MakeExpNumToStrNum(ByVal pNum)
   Dim tmpExpArr
   Dim tmpNumberArr
   Dim tmpZeroString
   Dim iPositive
   Dim tmpData
   
   If CDbl(pNum) > 0 Then
      iPositive = True
   Else
      iPositive = False
   End If
   
   tmpZeroString = "000000000000000000000000000"
   
   If InStr(pNum, "E+") Then
      tmpExpArr = Split(pNum, "E+")
      If InStr(tmpExpArr(0), ".") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ".")
         tmpData = tmpNumberArr(0) & Left(tmpNumberArr(1) & tmpZeroString, tmpExpArr(1))
      ElseIf InStr(tmpExpArr(0), ",") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ",")
         tmpData = tmpNumberArr(0) & Left(tmpNumberArr(1) & tmpZeroString, tmpExpArr(1))
      Else
         tmpData = tmpExpArr(0) & Left(tmpZeroString, tmpExpArr(1))
      End If
   ElseIf InStr(pNum, "E-") Then
      tmpExpArr = Split(pNum, "E-")
      If InStr(tmpExpArr(0), ".") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ".")
         tmpData = "0." & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpNumberArr(0) & tmpNumberArr(1)
      ElseIf InStr(tmpExpArr(0), ",") > 0 Then
         tmpNumberArr = Split(tmpExpArr(0), ",")
         tmpData = "0," & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpNumberArr(0) & tmpNumberArr(1)
      Else
         tmpData = "0." & Right(tmpZeroString, tmpExpArr(1) - 1) & tmpExpArr(0)
      End If
   Else
      tmpData = pNum
   End If
   
   If InStr(tmpData,"-") Then
      MakeExpNumToStrNum = "-" & Replace(tmpData,"-","")
   Else   
      MakeExpNumToStrNum = tmpData
   End If


End Function


%>