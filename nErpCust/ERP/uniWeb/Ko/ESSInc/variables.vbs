'========================================================================================
' Amount,Cost,Quantity,Exchange Rate Decimal Place Variable Definition
'========================================================================================
Class TB19029
     Dim DecPoint						'Decimal point
     Dim RndPolicy                      'Round Policy
     Dim RndUnit                        'Round Unit     
End Class

Dim ggAmtOfMoney						' Amount of Money
Set ggAmtOfMoney = New TB19029

Dim ggQty							' Quantity    
Set ggQty = New TB19029

Dim ggUnitCost						' Unit Cost
Set ggUnitCost = New TB19029

Dim ggExchRate						' Exchange Rate 
Set ggExchRate = New TB19029

DIm ggStrIntegeralPart                  ' Variable that contains value of  Integer Parts Places
DIm ggStrDeciPointPart                  ' Variable that contains value of  Decimal Parts Places

DIm ggStrMinPart                        ' Variable that contains minimum value
DIm ggStrMaxPart                        ' Variable that contains maximum value

'========================================================================================
'Global variable Reference
'========================================================================================

Dim gActionStatus
Dim gActiveElement       ' Variable that contains active element value in current Document
Dim gActiveSpdSheet      ' 2002/10/01  jslee
Dim gADODBConnString
Dim gADODBConnStringSA
Dim gAltNo
Dim gAPDateFormat        'AP Server DateFormat
Dim gAPDateSeperator     'AP Server DateSeperator
Dim gAPNum1000           'APServer Number 1000 unit seperator
Dim gAPNumDec            'APServer Number decimal pointer
Dim gAPServer
Dim gBConfMinorCD        ' gBConfMinorCD
Dim gBizArea
Dim gBizUnit
Dim gChangeOrgId
Dim gClientDateFormat    'Client DateFormat
Dim gClientDateSeperator 'Client DateSeperator
Dim gClientIp
Dim gClientNm
Dim gClientNum1000       'Client Number 1000 unit seperator
Dim gClientNumDec        'Client Number decimal pointer
Dim gComDateType         ' Company Date Delimiter Parameter
Dim gComNum1000          ' Company Number 1000 indication
Dim gComNumDec           ' Company Number Decimal Point indication
Dim gCompany 
Dim gCompanyNm
Dim gConnectionString
Dim gCostCenter
Dim gCountry 
Dim gCurrency
Dim gDatabase
Dim gDateFormat          ' Company Date format 
Dim gDateFormatYYYYMM    ' Company Date format 
Dim gDBServer 
Dim gDBServerNm
Dim gDBServerIP
Dim gDepart 
Dim gDsnNo
Dim gEbDbName            ' EASYBASE DB Name
Dim gEbEnginePath        ' EASYBASE Engine Path
Dim gEbEnginePath5        ' EASYBASE Engine Path
Dim gEbPkgRptPath        ' EASYBASE Package Report Path
Dim gEbPkgRptPath5        ' EASYBASE Package Report Path
Dim gEbUserName          ' EASYBASE User Name
Dim gEbUserPass          ' EASYBASE User Password
Dim gEbUsrRptPath        ' EASYBASE User Report Path
Dim gEbUsrRptPath5        ' EASYBASE User Report Path
Dim gEnvInf
Dim gFiscCnt 
Dim gFiscEnd 
Dim gFiscStart
Dim gFocusSkip           ' Variable that indicates "Don't Care about Focus of Current Document"
Dim gIm_Post_Flag
Dim gIntDeptCd
Dim gIsTab
Dim gLang
Dim gLocRndPolicy         'Round Rule For foreign currency money
Dim gLoginDt 
Dim gLogonGp 
Dim gPageNo
Dim gPlant
Dim gPlantNm  
Dim gPo_Post_Flag
Dim gPurGrp
Dim gPurOrg
Dim gSalesGrp
Dim gSalesOrg
Dim gSetupMod
Dim gSeverity
Dim gSo_Post_Flag
Dim gStorageLoc
Dim gTaxRndPolicy         'Round Rule For Calculating TAX
Dim gTabMaxCnt
Dim gToolBarBit
Dim gUserType
Dim gUsrEngName
DIm gUsrID 
DIm gUsrName
Dim gWorkCenter
Dim gEWare

'========================================================================================
'Global variable for numeric format  / added on 2001/11/28 
'========================================================================================
Dim gBDataType
Dim gBCurrency
Dim gBDecimals
Dim gBRoundingUnit
Dim gBRoundingPolicy

'========================================================================================
' row,column seperator for spreadsheet
'========================================================================================
Dim gRowSep
Dim gColSep

DIm gDBLoginID            'DBLoginID
Dim gDBLoginPwd           'DBLoginPwd
Dim gUsrNm                'User Name

' 2003-02-19 Kim In Tae
Dim gQMDPAlignOpt
Dim gIMDPAlignOpt

Dim gRdsUse
Dim gDBKind

Dim gUserIdKind            '2005-05-31