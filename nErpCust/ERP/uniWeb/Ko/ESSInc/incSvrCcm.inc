<%
'FarPoint Spread 6.0                   SPR32X60  : 41F841C1-AE16-11D5-8817-0050DA6EF5E5
'FarPoint Spread 6.0 (OLEDB)           fpSpr60   : 41F841C0-AE16-11D5-8817-0050DA6EF5E5
'FarPoint Spread 7.0 (OLEDB)           FPSPR70   : 7114683A-020D-4D16-80FD-6ACE384B66DF
'FarPoint Spread 7.0 (OLEDB) (UNICODE) fpSpru70  : 71146838-020D-4D16-80FD-6ACE384B66DF

'FarPoint Spread 6.0 Preview                    SPR32X60  : 41F841C7-AE16-11D5-8817-0050DA6EF5E5
'FarPoint Spread 6.0 Preview (OLEDB)            fpSpr60   : 41F841C6-AE16-11D5-8817-0050DA6EF5E5
'FarPoint Spread 7.0 Preview (OLEDB)            FPSPR70   : 71146846-020D-4D16-80FD-6ACE384B66DF
'FarPoint Spread 7.0 Preview (OLEDB) (UNICODE)  fpSpru70  : 71146844-020D-4D16-80FD-6ACE384B66DF

'========================================================================
' Operational mode
'========================================================================
Const OPMD_CMODE  = 1000      'Create Mode
Const OPMD_UMODE  = 1001      'Update Mode

'========================================================================
' Execution status message
'========================================================================
Const UID_M0001   = 1500      'Search
Const UID_M0002   = 1501      'Insert
Const UID_M0003   = 1502      'Delete
Const UID_M0004   = 1503      'Setup
Const UID_M0005   = 1504      'Update
Const UID_M0006   = 1505      'Batch

'==============================================================================
' server 
'==============================================================================
Const ggServerIP       = "localhost"
Const BizSize          = 20

'==============================================================================
' Class id for client UI
'==============================================================================

'Const gCLSIDFPSPD70      = """CLSID:71146838-020D-4D16-80FD-6ACE384B66DF"""
'Const gCLSIDFPVPSPD70    = """CLSID:71146844-020D-4D16-80FD-6ACE384B66DF"""

Const gCLSIDFPSPD      = """CLSID:41F841C1-AE16-11D5-8817-0050DA6EF5E5"""
Const gCLSIDFPVPSPD    = """CLSID:41F841C7-AE16-11D5-8817-0050DA6EF5E5"""

Const gCLSIDFPSPD39    = """CLSID:E2454650-4D87-11D2-B8B2-0000C00A958C"""
Const gCLSIDFPVPSPD39  = """CLSID:664E2200-24DB-11D2-9A82-444553540000"""

Const gCLSIDFPDT       = """CLSID:DD55D13A-EBF7-11D0-8810-0000C0E5948C"""
Const gCLSIDFPDS       = """CLSID:DD55D13D-EBF7-11D0-8810-0000C0E5948C"""

Const gCLSIDCFX        = """CLSID:21F49842-BFA9-11D2-A89C-00104B62BDDA"""

Const UNI2KTV_IDVER    = "CLASSID=""CLSID:3F110E9C-F48F-11D3-9734-00A0C96D8AA6"""

Const LR_SPACE_TYPE_00 = " CLASS=""BasicTB"" CELLSPACING=0 CELLPADDING=0"
Const LR_SPACE_TYPE_10 = " CLASS=""BasicTB"" CELLSPACING=0 CELLPADDING=0"
Const LR_SPACE_TYPE_20 = " CLASS=""BasicTB"" CELLSPACING=0"
Const LR_SPACE_TYPE_30 = " CLASS=""BasicTB"" CELLSPACING=0"
Const LR_SPACE_TYPE_40 = " CLASS=""BasicTB"" CELLSPACING=0"
Const LR_SPACE_TYPE_50 = " CLASS=""TB4"" CELLSPACING=0"
Const LR_SPACE_TYPE_60 = " CLASS=""TB3"" CELLSPACING=0"

Const HEIGHT_TYPE_00   = " HEIGHT=11"
Const HEIGHT_TYPE_01   = " HEIGHT=3"
Const HEIGHT_TYPE_02   = " HEIGHT=5"
Const HEIGHT_TYPE_03   = " HEIGHT=2"

'==============================================================================
' Ado query mode 
'==============================================================================
Const DISCONNUPD       = "1"                'Disconnect + Update Mode
Const DISCONNREAD      = "2"                'Disconnect + ReadOnly Mode

Const C_PARAMETER_WRITE = "N"

   Select Case UCase(Request.Cookies("unierp")("gLang"))
      Case "KO","TEMPLATE","TEMPLATE1"
                Response.CharSet = "euc-kr"                               'Korea
      Case "CN"
                Response.CharSet = "GB2312"                               'China
      Case "IR"
                Response.CharSet = "windows-1256"                            'Japan
      Case "JA"
                Response.CharSet = "shift_jis"                            'Japan
      Case "EN"
               'Response.CharSet = "windows-1252"                         'U.S.A
      Case "HU"
                Response.CharSet = "windows-1250"                         'Hungary
   End Select               

   If Request.Cookies("unierp")("gCharSet") = "U" Then
      Session.CodePage = 65001
      Response.CharSet = "utf-8"
      
   End If   
      

%>