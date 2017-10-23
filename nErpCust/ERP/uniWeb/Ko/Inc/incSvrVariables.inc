<%
Dim gAltNo    			 
Dim gADODBConnString      ' ADB DB Connection String
Dim gADODBConnStringSA    ' ADB DB Connection String
Dim gAPDateFormat         'AP Server DateFormat
Dim gAPDateSeperator      'AP Server DateSeperator
Dim gAPNum1000            'APServer Number 1000 unit seperator
Dim gAPNumDec             'APServer Number decimal pointer
Dim gAPServer             'AP Server
Dim gBConfMinorCD
Dim gBizArea              'Business Area
Dim gBizUnit              'Business Unit
Dim gChangeOrgId          'Organization Change id
Dim gClientDateFormat     'Client DateFormat
Dim gClientDateSeperator  'Client DateSeperator
Dim gClientIp             'Client IP
Dim gClientNm             'Client Computer Name
Dim gClientNum1000        'Client Number 1000 unit seperator
Dim gClientNumDec         'Client Number decimal pointer
Dim gConnectionString     'Comproxy Connection String
Dim gComDateType          'Date Delimeter    
Dim gComNum1000	          'Company 1000 Places Indication
Dim gComNumDec 			  'Company Decimal Point Indication
Dim gCompany              'Company
Dim gCompanyNm            'Company name
Dim gCostCenter           'CostCenter
Dim gCountry              'Country
Dim gCurrency             'Currency
Dim gDatabase             'Database
Dim gDateFormat           'Date Format
Dim gDateFormatYYYYMM     'Date Format YYYY ~ MM
DIm gDBLoginID            'DBLoginID
Dim gDBLoginPwd           'DBLoginPwd
Dim gDBServer             'DB Server
Dim gDBServerNm
Dim gDBServerIP
Dim gDepart               'Department
Dim gDsnNo                'DSN NO 
Dim gEnvInf
Dim gFiscCnt              'Fiscal Term
Dim gFiscEnd              'Ending Date of Fiscal Term
Dim gFiscStart            'Starting Date of Fiscal Term
Dim gIm_Post_Flag
Dim gIntDeptCd            'Internal Department Code
Dim gLang                 'Language
Dim gLocRndPolicy         'Round Rule For foreign currency money
Dim gLoginDt              'Login Date
Dim gLogoName
Dim gLogo
Dim gLogonGp              'Logon Group
Dim gPlant                'Plant
Dim gPlantNm              'Plant name
Dim gPo_Post_Flag
Dim gProgId               'Program ID
Dim gPurGrp               'Purchasing group
Dim gPurOrg               'Purchase Organization
Dim gSalesGrp             'Sales group
Dim gSetupMod             'Module id that installed
Dim gSalesOrg             'Sales Organization
Dim gSeverity             'Severity
Dim gSo_Post_Flag
Dim gStorageLoc           'Storage Location
Dim gStrGlobalCollection  ' Global string for component
Dim gTaxRndPolicy         'Round Rule For Calculating TAX
Dim gUsrEngName           'User Name(English)
Dim gUsrId                'User ID
Dim gUsrNm                'User Name
Dim gWorkCenter           'Work Center
Dim gEWare
Dim gEmpNo                'ESS
Dim gProAuth              'ESS
Dim gDeptAuth             'ESS
Dim gCanBeDebug           '2005-10-05
Dim gISOLLVL              '2005-10-05

Dim gUserIdKind            '2005-05-31

Dim intPwdFailCnt         'Count of Password Authentication error

Dim gUDF6                 'Count of Password Authentication error
Dim gUDF7                 'Count of Password Authentication error
Dim gUDF8                 'Count of Password Authentication error
Dim gUDF9                 'Count of Password Authentication error

'========================================================================================
' Structure for Message
'========================================================================================
Class FetchMsg
    Dim Severity
    Dim Text
    Dim CD
End Class

'==============================================================================
' row,column seperator for seperator
'==============================================================================
Dim gRowSep 
Dim gColSep 

Dim gDBKind              '2003-11-06

Dim adVarXChar           '2004/9/30 unicode ,sp, adVarChar
Dim adXChar              '2004/9/30 unicode ,sp, adChar


%>