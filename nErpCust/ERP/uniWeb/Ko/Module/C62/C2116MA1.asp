<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :  Ado query Sample with DBAgent(Multi + Multi)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2002/12/12
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>


<Script Language="VBScript">
Option Explicit                                                  '��: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c2116mb8.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DTL_ID = "c2116mb6.asp"			'��: �����Ͻ� ���� ASP�� 
const BIZ_PGM_BOM_ID = "c2116mb7.asp"	 '��: �����Ͻ� ���� ASP�� 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_SHEETMAXROWS_D_A  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

Const C_SHEETMAXROWS_D_B  = 5                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey            = 2                                    '�١١١�: Max key value

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

const C_ITEM_KEY = 15
const C_ITEM_FIELD = 1
const C_ITEM_CD = 1
const C_ITEM = 2
Const C_ITEM_ACCNT_UNT = 2
const C_ROLLUPAMT = 3	'�ι�° �ǿ����� ��� 

Const C_MAN_COST_M = 3    '�������� ���� 
const C_MAN_COST_L = 4    '�������� �빫�� 
const C_MAN_COST_E = 5    '�������� ��� 
const C_MAN_SUM = 6
const C_DI_COST_M = 7     '���� ���� 
const C_DI_COST_L = 8     '���� �빫�� 
const C_DI_COST_E = 9     '���� ��� 
const C_DI_SUM = 10
const C_IND_COST_M = 11     '���� ���� 
const C_IND_COST_L = 12     '���� �빫�� 
const C_IND_COST_E = 13     '���� ��� 
const C_IND_SUM = 14

Const C_Sep  = "/"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM = "PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"
Const C_SUBCON  = "SUBCON"

Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/ASSEMBLY.gif"
Const C_IMG_SUBCON = "../../../CShared/image/subcon.gif"


Const tvwChild = 4


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                             '��: Popup status                           
Dim gSelframeFlg											'��: Tab Flag
Dim lgPrevKey
Dim lgPrevKey2
'��:--------Spreadsheet #1-----------------------------------------------------------------------------   


Dim lgPageNo_A                                              '��: Next Key tag                          
Dim lgSortKey_A                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '��: Next Key tag                          
Dim lgSortKey_B                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet temp---------------------------------------------------------------------------   
                                                             '��:--------Buffer for Spreadsheet -----   

'Dim lgKeyPos                                                '��: Key��ġ                               
'Dim lgKeyPosVal                                             '��: Key��ġ Value                         

Dim IsOpenPop
'���� 
Dim lgSelNode


Dim BaseDate
BaseDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date
'  BaseDate     = Date(You must not code like this!!!!)                                       'Get AP Server Date


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode = Parent.OPMD_CMODE                         'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

	lgPrevKey = ""
	
	'���� 
	lgSelNode = ""
	
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

'--------------- ������ coding part(�������,Start)--------------------------------------------------

'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>                                '��: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("C2110MA101", "S", "A", "V20021212", Parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock ("A")

	Call SetZAdoSpreadSheet("C2110MA102", "S", "B", "V20021212", Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X", "X")
	Call SetSpreadLock ("B")

    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal iOpt )
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLock 1 , -1
          ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
            ggoSpread.SpreadLock 1, -1
            ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
'Sub PopSaveSpreadColumnInf()
'    ggoSpread.Source = gActiveSpdSheet
'    Call ggoSpread.SaveSpreadColumnInf()
'End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
'Sub PopRestoreSpreadColumnInf()
'    ggoSpread.Source = gActiveSpdSheet
'    Call ggoSpread.RestoreSpreadInf()
'    Call InitSpreadSheet()      
'   Call InitComboBox
'	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
'	Call initMinor()
'End Sub


'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029														'��: Load table , B_numeric_format
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
  
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                      ' ��: Lock  Suitable  Field
    

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal
	Call InitTreeImage	

	Call InitSpreadSheet()

    Call SetToolbar("1100000000001111")																		        '��: ��ư ���� ���� 
   	gTabMaxCnt = 2
	gIsTab = "Y"

    frm1.txtPlantCd.focus
    '--------- Developer Coding Part (End  ) ----------------------------------------------------------
    Set gActiveElement = document.activeElement 
End Sub

'========================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    Err.Clear     

	call ClickTab1
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
    
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								        '��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery("MQ") = False Then   
       Exit Function           
    End If     							

    FncQuery = True		
End Function


'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    FncInsertRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '��: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect) 
	Dim strVal
	
    Err.Clear                                                                    '��: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '��: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                               '��: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '��: Show Processing Message
	
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    Select Case pDirect
        Case "MQ","MN"
               ' Call CopyPopupInfABT("1")

                With Frm1
                If lgIntFlgMode = Parent.OPMD_CMODE Then
					strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'���� �˻��������� Query
					strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
					strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
					strVal = strVal & "&txtItemAccntCd=" & .txtItemAccntCd.value
					strVal = strVal & "&txtCItemCd=" & .txtCItemCd.value		
					strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows				'��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
					
	'--------- Developer Coding Part (End) ----------------------------------------------------------
                    strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          '��: Next key tag
                    strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
                    'strVal = strVal      & "&lgTailList="        & Parent.MakeSQLGroupOrderByList(UBound(lgFieldNM_T),lgPopUpR_T,lgFieldCD_T,lgNextSeq_T,lgTypeCD_T(0),Parent.C_MaxSelList)
                    'strVal = strVal      & "&lgSelectList="      & EnCoding(lgSelectList_A)
                     strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_A)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
					strVal = strVal      & "&lgPrevKey="        & lgPrevKey
				ELSE
					strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'���� �˻��������� Query
					strVal = strVal & "&txtPlantCd=" & .hPlantCd.value				
					strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
					strVal = strVal & "&txtItemAccntCd=" & .hItemAccntCd.value
					strVal = strVal & "&txtCItemCd=" & .hCItemCd.value		
					strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows				'��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
					
	'--------- Developer Coding Part (End) ----------------------------------------------------------
                    strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          '��: Next key tag
                    strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
                    'strVal = strVal      & "&lgTailList="        & Parent.MakeSQLGroupOrderByList(UBound(lgFieldNM_T),lgPopUpR_T,lgFieldCD_T,lgNextSeq_T,lgTypeCD_T(0),Parent.C_MaxSelList)
                    'strVal = strVal      & "&lgSelectList="      & EnCoding(lgSelectList_A)
                     strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_A)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
				 	 strVal = strVal      & "&lgPrevKey="        & lgPrevKey
				END IF
                End With
				
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    End Select		
    
    
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
	
    DbQuery = True                                                               '��: Processing is OK

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk( )											 '��: ��ȸ ������ ������� 
	
    lgIntFlgMode     = Parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
	
       'Call vspdData_Click(1,1)
       'frm1.vspdData.focus
	Call SetToolbar("1100000000011111")																		        '��: ��ư ���� ����	

	'Call ggoOper.LockField(Document, "Q")								 '��: This function lock the suitable field 
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================



'==========================================================================================
'   Function Name :LookUpBomNo
'   Function Desc :ǥ�ػ���ȸ Tab �� Ŭ���Ҷ� Ʈ�����·� BOM�� ��ȸ 
'==========================================================================================

Sub LookUpBomNo()
    
    Err.Clear															'��: Protect system from crashing
    
    Dim strVal
	Dim txtConFlg
	Dim lcSrchType

	frm1.txtHdnItemAcct.value = ""
	frm1.txtBomNo.value = ""
	
		
	IF LayerShowHide(1) = False Then
		Exit Sub
	END IF
	
	frm1.txtSrchType.value = "2"

	frm1.vspddata.col = C_ITEM_KEY
	frm1.vspddata.row = frm1.vspddata.activeRow

	
   
    	strVal = BIZ_PGM_BOM_ID & "?txtMode=" & Parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.vspdData.value)		'��: ��ȸ ���� ����Ÿ 
    
    	strVal = strVal & "&txtBaseDt="	&  BaseDate 
    	strVal = strVal & "&txtUpdtUserId=" & Parent.gUsrID
    	strVal = strval & "&rdoSrchType=" & Trim(frm1.txtSrchType.value)
    	strVal = strVal & "&txtBomNo=1"
	
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Sub

'==========================================================================================
'   Function Name :LookUpBomNoOk
'   Function Desc :������ ǰ���� BOM�� �����ϴ� �� üũ 
'==========================================================================================


Sub LookUpBomNoOk()
	Dim Node
	Dim PrntKey

    '----------------------------------------------------
      '- Parent Node�� Setting�ϰ� Header Data�� �����´�.
      '---------------------------------------------------
    	frm1.vspddata.col = C_ITEM_KEY
	frm1.vspddata.row = frm1.vspddata.activeRow   
	PrntKey = UCase(Trim(frm1.vspddata.value) & "|^|^|" & UCase(frm1.txtBomNo.value))

	If Trim(frm1.txtHdnItemAcct.value) = "10" Or Trim(frm1.txtHdnItemAcct.value) = "20" Then 
		Set Node = frm1.uniTree1.Nodes.Add(,,PrntKey,UCase(Trim(frm1.vspddata.Value)),C_PROD, C_PROD)      
		Node.Expanded = True
		Call SetFieldProp(0)
		Call SetModChange(0)												'BOM�� ���� ��츦 ���� ó�� ���¸� Header�Է»��·� 
	Else
		Exit Sub
	End If
	
	Set Node = Nothing
	
	
End Sub

'==========================================================================================
'   Function Name :LookUpBomNoNotOk
'   Function Desc :������ ǰ���� BOM�� �����ϴ� �� üũ 
'==========================================================================================


Sub LookUpBomNoNotOk()
	Dim Node
	Dim PrntKey
    '----------------------------------------------------
      '- Parent Node�� Setting�ϰ� Header Data�� �����´�.
      '---------------------------------------------------
    	frm1.vspddata.col = C_ITEM_KEY
	frm1.vspddata.row = frm1.vspddata.activeRow     
	PrntKey = UCase(Trim(frm1.vspddata.value) & "|^|^|" & UCase(frm1.txtBomNo.value))

	If Trim(frm1.txtHdnItemAcct.value) = "10" Or Trim(frm1.txtHdnItemAcct.value) = "20" Then 
		Set Node = frm1.uniTree1.Nodes.Add(,,PrntKey,UCase(Trim(frm1.vspddata.Value)),C_PROD, C_PROD)      
		Node.Expanded = True
		Call SetFieldProp(0)
		Call SetModChange(0)												'BOM�� ���� ��츦 ���� ó�� ���¸� Header�Է»��·� 
		'frm1.txtBOMDesc.focus 	
	Else
		Exit Sub
	End If
	
	Set Node = Nothing

End Sub

'===========================================================================
' Function Name : LookUpCostOfVspd() 
' Function Desc : Ʈ����� Ŭ���� ǰ���ڵ�� �������忡 �ִ� ����Ÿ �˻� 
'===========================================================================
Function LookUpCostOfVspd(byval pItemCd)
  dim i 
  Dim lSearchFlag 
  
  
  with frm1
    
    .txtItemUnt.value = ""
    .txtItemNmDesc.value = ""
        
	.txtDi_Mcost.text = 0
	.txtDi_Lcost.text = 0
	.txtDi_Ecost.text = 0
	.txtInd_Mcost.text = 0
	.txtInd_Lcost.text = 0
	.txtInd_Ecost.text = 0
	.txtInDi_Sum.text = 0
	.txtInInd_Sum.text = 0
	.txtDi_Sum.text  = 0
	.txtInd_Sum.text  = 0
	.txtInDi_Sum.text = 0
	.txtInInd_Sum.text = 0
	.txtOutDi_Sum.text = 0
	.txtOutInd_Sum.text = 0
	
        for i = 1 to .vspdData.maxrows
           .vspdData.row = i
           .vspdData.col = C_ITEM_KEY
           
           if .vspdData.value = pItemCd then
              lSearchFlag = i
              exit for
           end if
        next
        
        if lSearchFlag > 0 then  		
			.vspdData.row = lSearchFlag
			'.vspdData.col = C_ROW
			'.vspdData.row = cint(.vspdData.value)         'row ���� : active row �� ������ �࿣ ǰ��(3�پ�����) �� ������.
			         
			' ���� �ο쿡���� ���ο�����  
			.vspddata.col = C_DI_COST_M
			.txtDi_Mcost.text =  .vspddata.text						
			.vspddata.col = C_DI_COST_L 
			.txtDi_Lcost.text =  .vspddata.text
			.vspddata.col = C_DI_COST_E 
			.txtDi_Ecost.text =  .vspddata.text
			.vspddata.col = C_IND_COST_M 
			.txtInd_Mcost.text =  .vspddata.text
			.vspddata.col = C_IND_COST_L 
			.txtInd_Lcost.text =  .vspddata.text
			.vspddata.col = C_IND_COST_E
			.txtInd_Ecost.text =  .vspddata.text
		
			.vspddata.col = C_DI_SUM		'���ο��� ������ �� 
			.txtInDi_Sum.text =  .vspddata.text 

			.vspddata.col = C_IND_SUM		'���ο��� ������ �� 
			.txtInInd_Sum.text =  .vspddata.text 
	
			.vspdData.row = .vspdData.row + 1 ' �ܺ� ������ �÷��� 
			.vspdData.col = C_ITEM_FIELD            'ǰ����� 2��° �ٿ� �����Ƿ�   
			.txtItemNmDesc.value = .vspdData.text
			
			.vspdData.col = C_ITEM_ACCNT_UNT            '�԰��� 2��° �ٿ� �����Ƿ�   
			.txtItemUnt.value = .vspdData.text
		
			.vspddata.col = C_DI_SUM		'�ܺο��� ���� �� 
			.txtOutDi_Sum.text =  .vspddata.text
			
			.vspddata.col = C_IND_SUM		'�ܺο��� ���� �� 
			.txtOutInd_Sum.text =  .vspddata.text
			
			.vspddata.col = C_DI_COST_M
			.txtDi_Mcost.text =  UNIFormatNumber(UNICDbl(.txtDi_Mcost.text) + UNICDbl(.vspddata.text),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)
			
			.vspddata.col = C_DI_COST_L
			.txtDi_Lcost.text =  UNIFormatNumber(UNICDbl(.txtDi_Lcost.text) + UNICDbl(.vspddata.text),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)
			
			.vspddata.col = C_DI_COST_E 
			.txtDi_Ecost.text =  UNIFormatNumber((UNICDbl(.txtDi_Ecost.text) + UNICDbl(.vspddata.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)
		
			.vspddata.col = C_IND_COST_M 
			.txtInd_Mcost.text = UNIFormatNumber((UNICDbl(.txtInd_Mcost.text) + UNICDbl(.vspddata.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)
			
			.vspddata.col = C_IND_COST_L 
			.txtInd_Lcost.text =  UNIFormatNumber((UNICDbl(.txtInd_Lcost.text) + UNICDbl(.vspddata.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)			
			
			.vspddata.col = C_IND_COST_E
			.txtInd_Ecost.text =  UNIFormatNumber((UNICDbl(.txtInd_Ecost.text) + UNICDbl(.vspddata.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)			
		
			.txtDi_Sum.text   = UNIFormatNumber((UNICDbl(.txtDi_Ecost.text) + UNICDbl(.txtDi_Lcost.text) + UNICDbl(.txtDi_Mcost.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)			
			.txtInd_Sum.text  = UNIFormatNumber((UNICDbl(.txtInd_Ecost.text) + UNICDbl(.txtInd_Lcost.text) + UNICDbl(.txtInd_Mcost.text)),ggUnitCost.Decpoint,-2,0,ggUnitCost.RndPolicy,ggUnitCost.RndUnit)			


		end if
    end with	
     
end function



'========================================================================================
' Function Name : InitTreeImage
' Function Desc : �̹��� �ʱ�ȭ 
'========================================================================================

Function InitTreeImage()
	Dim NodX, lHwnd
	
	With frm1

	.uniTree1.SetAddImageCount = 4
	.uniTree1.Indentation = "200"	' �� ���� 
	.uniTree1.AddImage C_IMG_PROD, C_PROD, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_MATL, C_MATL, 0
	.uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
	.uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0

	.uniTree1.OLEDragMode = 0														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
	.uniTree1.OLEDropMode = 0
	
	End With

End Function


'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click�� Look Up Call
'==========================================================================================


Sub uniTree1_NodeClick(ByVal Node)
    Dim strVal
    
    Dim NodX
    

	Dim iPos2

	Dim txtItemCd
  

	Err.Clear                                                               '��: Protect system from crashing
			
	frm1.vspdData2.maxrows = 0
	
   	With frm1
	
    Set NodX = .uniTree1.SelectedItem

   
    
    If Not NodX Is Nothing Then				' ���õ� ������ ������ 

		'-------------------------------------
		'If Same Node Clicked, Exit
		'---------------------------------------
			
		If NodX.Key = lgSelNode Then
			Set NodX = Nothing
			Exit Sub
		Else
			lgSelNode = NodX.Key
		End If



		iPos2 = InStr(NodX.Text, "    (")   
		txtItemCd = Trim(Left(NodX.Text,iPos2-1))

		
		IF LayerShowHide(1) = False Then
			Exit Sub
		END IF
 	
		strVal = BIZ_PGM_DTL_ID & "?txtMode=" & Parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(txtItemCd)					'��: LookUP ���� ����Ÿ 
		strVal = strVal & "&lgPrevKey2=" & lgPrevKey2
        strVal = strVal & "&lgSelectListDT="    & GetSQLSelectListDataType("B")
		strVal = strVal & "&lgMaxCount="    & CStr(C_SHEETMAXROWS_D_A)
				
		Call RunMyBizASP(MyBizASP, strVal)	
	
	
	End If
    
    Set NodX = Nothing
    
    End With
    


End Sub







'===========================================================================
' Function Name : OpenMinor
' Function Desc : OpenMinor Reference Popup
'===========================================================================
Function OpenMinor(ByVal iMinor)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim itemacct

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iMinor
	Case 0												
		arrParam(0) = "ǰ������˾�"				' �˾� ��Ī 
		arrParam(1) = "B_MINOR a,b_item_acct_inf b"							' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtItemAccntCd.value)		' Code Condition
		arrParam(3) = ""	' Name Cindition
		arrParam(4) = "a.MAJOR_CD=" & FilterVar("P1001", "''", "S") & " and a.minor_cd = b.item_acct and b.item_acct_group <> " & FilterVar("6MRO","''","S")	 			' Where Condition
		arrParam(5) = "ǰ�����"						' TextBox ��Ī 
		
	    arrField(0) = "MINOR_CD"						' Field��(0)
	    arrField(1) = "MINOR_NM"						' Field��(1)
	    
	    arrHeader(0) = "ǰ������ڵ�"					' Header��(0)
	    arrHeader(1) = "ǰ�������"						' Header��(1)

	Case 1		
		arrParam(0) = "ǰ���˾�"				' �˾� ��Ī 
		arrParam(1) = "b_item a,b_item_by_plant b"							' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtCItemCd.value)		' Code Condition
		arrParam(3) = ""								' Name Cindition
		
		itemacct = Trim(frm1.txtItemAccntCd.value)
		IF itemacct = "" Then
				 itemacct = "%"
		END If
	
		arrParam(4) = "a.item_cd = b.item_cd and b.plant_cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "" _
			& " and a.valid_flg = " & FilterVar("y", "''", "S") & "  and a.valid_from_dt <= getdate() and a.valid_to_dt >= getdate() " _
			& " and b.item_acct LIKE  " & FilterVar(itemacct, "''", "S") & ""
		
		arrParam(5) = "ǰ��"						' TextBox ��Ī 
		
	    arrField(0) = "a.item_cd"						' Field��(0)
	    arrField(1) = "a.item_nm"						' Field��(1)
	    
	    arrHeader(0) = "ǰ���ڵ�"					' Header��(0)
	    arrHeader(1) = "ǰ���"						' Header��(1)


	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinor(arrRet,iMinor)
	End If	
End Function



'======================================================================================================
'	Name : OpenPlant()
'	Description : Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

 '------------------------------------------  SetMinor()  --------------------------------------------------
'	Name : SetMinor()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetMinor(Byval arrRet,ByVal iMinor)

If arrRet(0) <> "" Then 
	Select Case iMinor
	Case 0												' ���� 
		frm1.txtItemAccntCd.value = arrRet(0)
		frm1.txtItemAccntNm.value = arrRet(1)
	Case 1												' ���� 
		frm1.txtCItemCd.value = arrRet(0)
		frm1.txtCItemNm.value = arrRet(1)
		
	end select
End If

End Function


'======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
			
End Function


Function OpenPopUp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("125000","x","x","x") '������ ���� �Է��ϼ��� 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtCItemCd.value)	' Item Code
	arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field��(0) :"ITEM_CD"
	arrField(1) = 2									' Field��(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/B1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet)
	End If	

End Function

 '==========================================  2.4.3 SetPopup()  =============================================
'	Name : SetPopup()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function SetPopUp(Byval arrRet)
	With frm1
		.TxtCItemCd.Value = arrRet(0)
		.TxtCItemNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
		
	End With
	
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"                  
	            gPos = "B"
   End Select     
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


 '==========================================  2.3.1 Tab Click ó��  =================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tabó�� �κ� (Header Tab�� �ִ� ��츸 ���)  ---------------------------- 
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	
	
	 '++++++++++++  Insert Your Code  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
	'Call SetToolBar(pstr)
   	 ' ----��Ŀ�� �̵� --- 
   	'Call setFocus(CLICK_HEADER)
	 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
End Function

Function ClickTab2()
     
     
    
	If gSelframeFlg = TAB2 Then Exit Function
	
    if frm1.vspdData.maxrows = 0 then Exit Function 

'	frm1.vspdData2.maxRows = 0
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	
	Call changeTabs(TAB2)	 '~~~ �ι�° Tab 
	gSelframeFlg = TAB2
	 '++++++++++++  Insert Your Code  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
											
    													  
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
	frm1.vspddata.col = C_ITEM_KEY
	frm1.vspddata.row = frm1.vspddata.activerow
	call LookUpCostOfVspd(frm1.vspddata.value)
	
	frm1.uniTree1.Nodes.Clear
	'Call ggoOper.ClearField(Document, "2")

    call LookUpBomNo
    						
    
   	 ' ----��Ŀ�� �̵� --- 
   	'Call setFocus(CLICK_HEADER)
	 '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 '   Dim ii

'	Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If
    
  	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
'	 For ii = 1 to UBound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii)   = frm1.vspdData.text
'	 Next
	 
     frm1.vspdData2.MaxRows = 0
     lgPageNo_B       = ""                                  'initializes Previous Key
     lgSortKey_B      = 1

'--------------- ������ coding part(�������,Start)----------------------------------------------------
     'Call DbQuery("M1Q")
'--------------- ������ coding part(�������,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData2_Click( ByVal Col, ByVal Row)
'	Call SetPopupMenuItemInf("00000000001") 
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	'----------------------
	'Column Split
	'----------------------
'--------------- ������ coding part(�������,End)------------------------------------------------------
    Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData2
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgPageNo_A <> "" and lgPrevKey <> "" Then                            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'��: ������ üũ'
		If lgPageNo_B <> "" and lgPrevKey2 <> "" Then                            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("M1N") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
    
End Sub

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End if
End Sub
'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : fpdtFromEnterDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub fpdtFromEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'=======================================================================================================
'   Event Name : fpdtToEnterDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub fpdtToEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	
'======================================================================================================= -->


<BODY TABINDEX="-1" SCROLL="No">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǥ�ؿ�����ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǥ�ؿ��� ����ȸ</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">����</TD>
									<TD CLASS="TD6"><INPUT CLASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
														<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="�����" tag="14X"></TD>
										
									<TD CLASS="TD5">ǰ�����</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemAccntCD" MAXLENGTH="2" SIZE=10  ALT ="ǰ�����" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAccntCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenMinor(0)">
														<INPUT NAME="txtItemAccntNM" MAXLENGTH="30" SIZE=20  ALT ="ǰ���" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">ǰ��</TD>
									<TD CLASS="TD6"><INPUT  NAME="txtCItemCD" MAXLENGTH="18" SIZE=10  ALT ="ǰ��" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup()">
														<INPUT NAME="txtCItemNM" MAXLENGTH="30" SIZE=25  ALT ="ǰ���" tag="14X"></TD>
										
								    <TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>

						<!-- ù��° �� ����  -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
								</TR>
							</TABLE>
						</DIV>

						<!-- �ι�° �� ����  -->
						<DIV ID="TabDiv"  SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
							    <TR>
									<TD HEIGHT=100% WIDTH=40%>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=uniTree1 width=100% height=100% <%=UNI2KTV_IDVER%>> <PARAM NAME="ImageWidth" VALUE="16">  <PARAM NAME="ImageHeight" VALUE="16">  <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7">  <PARAM NAME="LabelEdit" VALUE="1">  </OBJECT>');</SCRIPT>
									</TD>
								    <TD HEIGHT=100% WIDTH=60%>
										<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 12%">
									        <TABLE CLASS="BasicTB" CELLSPACING=0>
										        <TR>
												    <TD CLASS=TD5 NOWRAP>ǰ���</TD>
												    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNmDesc" SIZE=30  tag="24" ALT="ǰ���"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>�԰�</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemUnt" SIZE=30  tag="24" ALT="�԰�"></TD>
							                    </TR>  
									        </TABLE>
										</FIELDSET>        
											
										<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 24%">        							        
											<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=2>
											    <TR>
												    <TD CLASS=TD5 NOWRAP></TD>
													<TD CLASS=TD6 NOWRAP>����</TD>
													<TD CLASS=TD6 NOWRAP>����</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>����</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDi_Mcost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="��������" tag="24X4" id=OBJECT1> </OBJECT>');</SCRIPT>
													</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInd_Mcost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="��������" tag="24X4" id=OBJECT2> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>�빫��</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDi_Lcost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����빫��" tag="24X4" id=OBJECT3> </OBJECT>');</SCRIPT>
													</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInd_Lcost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����빫��" tag="24X4" id=OBJECT4> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>���</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDi_Ecost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="�������" id=OBJECT5> </OBJECT>');</SCRIPT>
													</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInd_Ecost style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="�������" id=OBJECT6> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
										    </TABLE>	
										 </FIELDSET>	    
											 
										 <FIELDSET CLASS="CLSFLD">			
									        <TABLE CELLSPACING=0 CELLPADDING=4 WIDTH="100%">
												<TR>
													<TD CLASS=TD5 NOWRAP>��</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDi_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="������ ��" id=OBJECT7> </OBJECT>');</SCRIPT>
													</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInd_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="������ ��" id=OBJECT8> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>���ο���</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInDi_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="������" id=OBJECT9> </OBJECT>');</SCRIPT>
													</TD>
							                        <TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtInInd_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="������" id=OBJECT10> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP HEIGHT=10>�ܺο���</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtOutDi_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="�ܺο��� ����" id=OBJECT11> </OBJECT>');</SCRIPT>
													</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtOutInd_Sum style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X4" ALT="�ܺο��� ����" id=OBJECT12> </OBJECT>');</SCRIPT>
													</TD>													
												</TR>
											</table>
										</FIELDSET>	    
											 
										<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 42%">	
											<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
												<TR>
													<TD HEIGHT=150>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
												</TR>
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtHdnItemAcct" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAccntCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCItemCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtSrchType" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="TxtItemNm" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="TxtItemNm1" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="TxtBOMNo" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="TxtBOMNo1" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtBOMDesc" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtItemCd1" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtItemAcct" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtItemAcctNm" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtDrawNo" tag="24" TABINDEX= "-1">
<% '�߰� For p1401mb10.asp %>
<INPUT TYPE=HIDDEN NAME="txtSpec" tag="24" TABINDEX= "-1">

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemFromDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" MAXLENGTH="10" SIZE="10" VIEWASTEXT id=fpDateTime1></OBJECT>');</SCRIPT>
							&nbsp;~&nbsp;
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemToDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" MAXLENGTH="10" SIZE="10" VIEWASTEXT id=fpDateTime2></OBJECT>');</SCRIPT>

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="" name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24X1" ALT="��ȿ�Ⱓ" MAXLENGTH="10" SIZE="10" VIEWASTEXT> </OBJECT>');</SCRIPT>
&nbsp;~&nbsp;
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="" name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="��ȿ�Ⱓ" tag="24X1" VIEWASTEXT> </OBJECT>');</SCRIPT>
<INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoDefaultFlg1">
<INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg2" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoDefaultFlg1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
