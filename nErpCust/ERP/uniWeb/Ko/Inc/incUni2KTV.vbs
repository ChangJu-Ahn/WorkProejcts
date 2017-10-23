
'-----------------------------------------------------------------------------------
' TreeLine Constant
'-----------------------------------------------------------------------------------
Const tvwTreeLines = 0 							'Show Treelines
Const tvwRootLines = 1 							'Show Rootlines in addition to Treelines

'-----------------------------------------------------------------------------------
'TreeRelationship Constant
'-----------------------------------------------------------------------------------
Const tvwFirst     = 0 							'First 사이블링. 
Const tvwLast      = 1 							'Last 사이블링. 
Const tvwNext      = 2 							'Next 사이블링. 
Const tvwPrevious  = 3 							'Previous 사이블링. 
Const tvwChild     = 4 							'Child structure

'-----------------------------------------------------------------------------------
'TreeStyle Constant
'-----------------------------------------------------------------------------------
Const tvwTextOnly 						= 0		'Text only
Const tvwPictureText 					= 1		'Picture & Text
Const tvwPlusMinusText 					= 2		'+/- & Text
Const tvwPlusPictureText 				= 3		'+/-, Picture & Text
Const tvwTreelinesText 					= 4 	'Treelines & Text
Const tvwTreelinesPictureText 			= 5 	'Treelines & Picture & Text
Const tvwTreelinesPlusMinusText 		= 6 	'Treelines, +/-, Text
Const tvwTreelinesPlusMinusPictureText 	= 7 	'Treelines, +/-, Picture & Text

'-----------------------------------------------------------------------------------
'LabelEdit Constant
'-----------------------------------------------------------------------------------
Const tvwAutomatic = 0 							'Indicating that label Editing is Automatic
Const tvwManual    = 1 							'calls label Editing

'-----------------------------------------------------------------------------------
'DragDrop Constant
'-----------------------------------------------------------------------------------
Const vbDropEffectNone =  0
Const ccNoDrop         = 12
Const ccDefault        =  0

'-----------------------------------------------------------------------------------
'ImageList의 Image Key Constant User Definition
'-----------------------------------------------------------------------------------
' shjshj
Const C_Open    = "Open"
Const C_Folder  = "Folder"
Const C_URL     = "URL"
Const C_None    = "None"
Const C_Const   = "Const"

'khy
Const C_USOpen    = "USOpen"
Const C_USFolder  = "USFolder"
Const C_USURL     = "USURL"
Const C_USNone    = "USNone"
Const C_USConst   = "USConst"
 
' by Shin hyoung jae 2001/3/6

Const C_AC    = "Account Close"
Const C_BC    = "Base Close"
Const C_CC    = "Cost Close"
Const C_DC    = "Configuration Close"
Const C_GC    = "Profit&Loss Close"
Const C_HC    = "HumanResource Close"
Const C_IC    = "Inventory Close"
Const C_JC    = "Eis Close"
Const C_MC    = "Purchase Close"
Const C_OC    = "Executive Close"
Const C_PC    = "Product Close"
Const C_QC    = "Quality Close"
Const C_RC    = "Process Close"
Const C_SC    = "Sales Close"
Const C_UC    = "User Close"
Const C_ZC    = "System Close"

Const C_AO    = "Account Open"
Const C_BO    = "Base Open"
Const C_CO    = "Cost Open"
Const C_DO    = "Configuration Open"
Const C_GO    = "Profit&Loss Open"
Const C_HO    = "HumanResource Open"
Const C_IO    = "Inventory Open"
Const C_JO    = "Eis Open"
Const C_MO    = "Purchase Open"
Const C_OO    = "Executive Open"
Const C_PO    = "Product Open"
Const C_QO    = "Quality Open"
Const C_RO    = "Process Open"
Const C_SO    = "Sales Open"
Const C_UO    = "User Open"
Const C_ZO    = "System Open"

'-----------------------------------------------------------------------------------
'uni2K Treeview의 Menu Index Constant
'-----------------------------------------------------------------------------------
Const C_MNU_OPEN   = 0
Const C_MNU_ADD    = 1
Const C_MNU_DELETE = 2
Const C_MNU_RENAME = 3