'========================================================================
'  File Name       : Common Constant Module(CCM)
'  description     : Define Constants that are used in common in Business Program Each
'  (Begin)Made by  :
'  (End) Made by   :
'  (Begin)Made on  : Oct/13th/1999
'  (End) Made on   : 
'  Update History  :
'  Comment         : Developer in charge of Business Program isn't allowed to update this Module Program
'========================================================================

'========================================================================
' Operational mode
'========================================================================
Const OPMD_CMODE         = 1000        'Create Mode
Const OPMD_UMODE         = 1001        'Update Mode

'========================================================================
' Message Reference Constants during Execution
'========================================================================
Const UID_M0001         = 1500         'Search
Const UID_M0002         = 1501         'Insert
Const UID_M0003         = 1502         'Delete
Const UID_M0004         = 1503         'setup
Const UID_M0005         = 1504         'Update
Const UID_M0006         = 1505         'Batch

'==============================================================================
'
'==============================================================================
Const TBC_QUERY         = "1001" 
Const TBC_NEW           = "1002" 
Const TBC_DELETE        = "1003" 
Const TBC_SAVE          = "1004" 
Const TBC_INSERTROW     = "1005" 
Const TBC_DELETEROW     = "1006" 
Const TBC_CANCEL        = "1007" 
Const TBC_PREV          = "1008" 
Const TBC_NEXT          = "1009" 
Const TBC_COPYRECORD    = "1010" 
Const TBC_EXPORT        = "1011" 
Const TBC_PRINT         = "1012" 
Const TBC_FIND          = "1013" 
Const TBC_HELP          = "1014" 
Const TBC_EXIT          = "1015" 

'==============================================================================
' Define User Defined Color
'==============================================================================
Const UC_REQUIRED_BAK   = &H99F7FF     'Color representing that Space should be Essentially input
Const UC_REQUIRED       = &HB4FFFF     'Color representing that Space should be Essentially input
Const UC_PROTECTED      = &Hdddddd     'Color representing that Space can not be input
Const UC_DEFAULT        = &HFFFFFF     'Color representing that Space can optionally be input

Const UCN_REQUIRED      = "required"   'Required  field
Const UCN_PROTECTED     = "protected"  'Protected field
Const UCN_DEFAULT       = "normal"     'Optional  field

'==============================================================================
' Class id for multi grid of uniSIMS 
'==============================================================================
Const UCN_GRID_TITLE    = "grid_title" 'Girid Tilte		'uniSIMS 
Const UCN_GPROTECTED    = "gprotected" 'Protected field		'uniSIMS
Const UCN_GREQUIRED     = "grequired"  'Required field		'uniSIMS
Const UCN_TPROTECTED    = "tprotected" 'title Protected field	'uniSIMS

'==============================================================================
' Message Box Reference Constants 
'==============================================================================
Const VB_YES_NO         = 36           '
Const VB_INFORMATION    = 64           '

'==============================================================================
' Cool:Gen Message Reference Constants 
'==============================================================================
Const MSG_OK_STR        = "990000"

'==============================================================================
' Date Format Reference Constants 
'==============================================================================
Const gServerDateFormat = "YYYY-MM-DD"	' Server Date format(Fixed)
Const gServerDateType   = "-"           ' Server Date Delimiter Parameter format(Fixed)
Const gServerBaseDate   = "1900-01-01"
Const gCommMaximumDate  = "2999-12-31"  ' 2002/10/12 lee jinsoo

'==============================================================================
' Window Screen Reference Constants for Find Window
'==============================================================================
Const C_SINGLE          = 0        ' Single Window Screen
Const C_MULTI           = 1        ' Multi Window Screen
Const C_SINGLEMULTI     = 2        ' Single/Multi Window Screen

'==============================================================================
' ADO Template  Reference Constants 
'==============================================================================
Const C_MaxSelList      = 6        ' Groupby or Sort list Maximum for Data
Const DISCONNUPD        = "1"	   ' Disconnect + Update Mode
Const DISCONNREAD       = "2"	   ' Disconnect + ReadOnly Mode

'==============================================================================
'  Numeric Format Information  Reference Constants in Master Data
'==============================================================================
Const ggAmtExOfMoneyNo  = "1"      ' Amount No
Const ggAmtOfMoneyNo    = "2"      ' Amount No
Const ggQtyNo           = "3"      ' Quantity No
Const ggUnitCostNo      = "4"      ' Cost No
Const ggExchRateNo      = "5"      ' Exchange Rate No

'==============================================================================
' Conversion rule no for Rnd Policy
'==============================================================================
Const gTaxRndPolicyNo   = "1" 
Const gLocRndPolicyNo   = "2" 
'==============================================================================
' SpreadSheet Cell Type Reference Constants
'==============================================================================
Const SS_CELL_TYPE_FLOAT = 2

'==============================================================================
' Aggregation
'==============================================================================
Const C_RGB_Sub_Total   = 13565902
Const C_RGB_Total       = 13565951
Const C_RGB_Grand_Total = 16770765

'==============================================================================
' SpreadSheet Action Reference Constants
'==============================================================================
Const SS_SCROLLBAR_NONE     = 0         ' Does not display scroll bars	SS_SCROLLBAR_NONE
Const SS_SCROLLBAR_H_ONLY   = 1         ' Displays horizontal scroll bar	SS_SCROLLBAR_H_ONLY
Const SS_SCROLLBAR_V_ONLY   = 2         ' Displays vertical scroll bar	SS_SCROLLBAR_V_ONLY
Const SS_SCROLLBAR_BOTH     = 3         ' (Default) Displays horizontal and vertical scroll bars	SS_SCROLLBAR_BOTH

Const SS_ACTION_ACTIVE_CELL = 0         ' Sets the active cell
      

'==============================================================================
' CELL TYPE    2002-10-22
'==============================================================================
Const CT_DATE        =  0  ' Date Creates date cell   
Const CT_EDIT        =  1  ' Edit (Default) Creates edit cell   
Const CT_FLOAT       =  2  ' FLOAT cell   
Const CT_INTEGER     =  3  ' INTEGER cell   
Const CT_PIC         =  4  ' PIC Creates PIC cell   
Const CT_STATIC_TEXT =  5  ' Static Text Creates static text cell   
Const CT_TIME        =  6  ' Time Creates time cell   
Const CT_BUTTON      =  7  ' Button Creates button cell   
Const CT_COMBOBOX    =  8  ' Combo Box Creates combo box cell   
Const CT_PICTURE     =  9  ' Picture Creates picture cell   
Const CT_CHECKBOX    = 10  ' Check Box Creates check box cell   
Const CT_OWNER_DRAWN = 11  ' Owner-Drawn Creates owner-drawn cell   
Const CT_CURRENCY    = 12  ' Currency Creates currency cell   [6.0]
Const CT_NUMBER      = 13  ' Number Creates numeric cell      [6.0]
Const CT_PERCENT     = 14  ' Percent Creates percent cell     [6.0]

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' 2002-11-11 
Const gAllowDragDropSpread = "T"
Const gForbidDragDropSpread = "S"

Const C_SORT_DBAGENT  = "QSDBAGENT"
Const C_GROUP_DBAGENT = "QGDBAGENT"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Const SORTW_WIDTH     = "675"
Const SORTW_HEIGHT    = "500"
Const GROUPW_HEIGHT   = "500"
Const GROUPW_WIDTH    = "675"

'==============================================================================
'
'==============================================================================
Const C_CHUNK_ARRAY_COUNT = 200
Const C_FORM_LIMIT_BYTE   = 102399
