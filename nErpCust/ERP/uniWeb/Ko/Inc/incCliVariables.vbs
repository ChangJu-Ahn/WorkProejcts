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
'Global variable for numeric format  / added on 2001/11/28 
'========================================================================================
Dim gBDataType
Dim gBCurrency
Dim gBDecimals
Dim gBRoundingUnit
Dim gBRoundingPolicy

'========================================================================================
Dim gActiveElement 
Dim gActionStatus
Dim gActiveSpdSheet      ' 2002/10/01  jslee
'========================================================================================
Dim gPageNo
Dim gIsTab
Dim gTabMaxCnt
Dim gFocusSkip     
Dim gPopupMenuItemBitInf
Dim PopupParent	

'========================================================================================
Dim gEnvInf
Dim gToolBarBit
Dim gADODBConnString
Dim gDsnNo
Dim gDateFormat          ' Company Date format 
Dim gLang

'========================================================================================
Dim gServerIP
Dim gLogoName
Dim gLogo

Dim gRdsUse

Dim gCharSet
Dim gCharSQLSet
