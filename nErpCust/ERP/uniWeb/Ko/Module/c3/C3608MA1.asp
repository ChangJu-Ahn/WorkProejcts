
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Cost
*  2. Function Name        : 표준원가대비실제원가조회 
*  3. Program ID           : C3608MA1
*  4. Program Name         : 표준원가대비실제원가조회 
*  5. Program Desc         : 표준원가대비실제원가조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/06/14
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Lee Tae Soo
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c3608mb8.asp"				'☆: 비지니스 로직 ASP명 
Const BIZ_LOOKUP_ID = "c3608mb9.asp"			'☆: 비지니스 로직 ASP명 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 
const C_ITEM_CD = 1
const C_ITEM_NM = 2
const C_ItemAcct	= 3

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          


Dim lgMaxFieldCount

Dim lgSaveRow 


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
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim StartDate
	StartDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

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

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub



'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
'Call SetCombo(frm1.cboPrcFlg, "T", "진단가")

'Call SetCombo(frm1.cboPrcFlg, "F", "가단가")
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
			
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    
	Call SetZAdoSpreadSheet("C3608MA101","S","A","V20021214",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock() 

End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
'	Call initMinor()
End Sub


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

    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
'    Call Parent.GetAdoFieldInf("C3608MA101","S","A")                                ' S for Sort , A for SpreadSheet No('A','B',....             

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   

'    lgMaxFieldCount =  UBound(Parent.gFieldNM)                      

'    ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)

'    Call Parent.MakePopData(Parent.gDefaultT,Parent.gFieldNM,Parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,Parent.C_MaxSelList)

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")
    Call InitComboBox()
    frm1.txtYyyymm.focus
    
    frm1.txtTotQty.allownull = False
    
    frm1.txtStd_Mcost1.allownull = False
    frm1.txtReal_Mcost1.allownull = False
    frm1.txtDiff_Mcost1.allownull = False
    
    frm1.txtStd_Lcost1.allownull = False
    frm1.txtReal_Lcost1.allownull = False
    frm1.txtDiff_Lcost1.allownull = False
    
    frm1.txtStd_Ecost1.allownull = False
    frm1.txtReal_Ecost1.allownull = False
    frm1.txtDiff_Ecost1.allownull = False
    
    frm1.txtStd_Sum1.allownull = False
    frm1.txtReal_Sum1.allownull = False
    frm1.txtDiff_Sum1.allownull = False
    
    frm1.txtStd_Mcost2.allownull = False
    frm1.txtReal_Mcost2.allownull = False
    frm1.txtDiff_Mcost2.allownull = False
    
    frm1.txtStd_Lcost2.allownull = False
    frm1.txtReal_Lcost2.allownull = False
    frm1.txtDiff_Lcost2.allownull = False
    
    frm1.txtStd_Ecost2.allownull = False
    frm1.txtReal_Ecost2.allownull = False
    frm1.txtDiff_Ecost2.allownull = False
    
    frm1.txtStd_Sum2.allownull = False
    frm1.txtReal_Sum2.allownull = False
    frm1.txtDiff_Sum2.allownull = False
    
    Set gActiveElement = document.activeElement	
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    						
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------

    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
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
Function DbQuery() 
	Dim strVal
	Dim strYear, strMonth, strDay


    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
   	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    With frm1

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  = Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'Hidden의 검색조건으로 Query
			strVal = strVal & "&txtYyyymm=" & strYear & strMonth
			strVal = strVal & "&txtPlantCd=" &  .hPlantCd.value				
			strVal = strVal & "&txtFrItemCd=" &  .hFrItemCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'현재 검색조건으로 Query
			strVal = strVal & "&txtYyyymm=" & strYear & strMonth
			strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
			strVal = strVal & "&txtFrItemCd=" & .txtFrItemCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
'       strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
        
        Call RunMyBizASP(MyBizASP, strVal)							
        
    
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
	
    Call SetToolbar("11000000000111")
    
    Call DbQuery2


End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery2() 
    
	Dim strVal
	Dim strYear, strMonth, strDay

	Err.Clear                                                               			'☜: Protect system from crashing
	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	DbQuery2 = False

	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    With frm1
    
    .vspdData.Row = .vspdData.ActiveRow
    .vspdData.Col = C_Item_CD
    
     '@Query_Text     
	strVal = BIZ_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001						'현재 검색조건으로 Query
	strVal = strVal & "&txtyyyymm=" &  strYear & strMonth
	strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
	strVal = strVal & "&txtItemCd=" &  .vspdData.Value
	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery2 = True
    
End Function

'======================================================================================================
' Function Name : DbQuery2Ok
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQuery2Ok()													'조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
	
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'======================================================================================================
'	Name : OpenCostCd()
'	Description : Cost Center PopUp
'=======================================================================================================
Function OpenCostCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "코스트센타팝업"			'팝업 명칭 
	arrParam(1) = "B_COST_CENTER"						'TABLE 명칭 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Condition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "코스트센타"			
	
    arrField(0) = "COST_CD"					    'Field명(0)
    arrField(1) = "COST_NM"					    'Field명(1)
    
    arrHeader(0) = "코스트센타코드"					'Header명(0)
    arrHeader(1) = "코스트센타명"					'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCd(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCostCd()
'	Description : Cost Center Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCostCd(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		.txtCostCd.value = arrRet(0)
    		.txtCostNm.value = arrRet(1)
    	Else
    		.vspdData.Col = C_CostCd
    		.vspdData.Text = arrRet(0)
    		.vspdData.Col = C_CostNm
    		.vspdData.Text = arrRet(1)
            
    		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		        '변경이 읽어났다고 알려줌 
    	End If
	
	End With
	
End Function

'======================================================================================================
'	Name : OpenPlant()
'	Description : Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCD.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function


'======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCD.focus
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
			
End Function


Function OpenItemCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtFrItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet, iWhere)
	End If	

End Function

Function SetItemCd(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		.txtFrItemCd.focus
    		.txtFrItemCd.value = arrRet(0)
    		.txtFrItemNm.value = arrRet(1)
    	End If
	
	End With
	
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")	   
	   Exit Function        
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
    End If
End Sub


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
		
    Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then                                             'If there is no data.
		Exit Sub
	End If	
	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	
'	For ii = 1 to Ubound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii)   = frm1.vspdData.text
		
'	Next
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    Dbquery2
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
    
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub  
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>표준원가대비실제원가조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% colspan=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=20 colspan=2>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/c3608ma1_fpDateTime1_txtYyyymm.js'></script>
									</TD>	
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6"><INPUT  ClASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
											<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=30  ALT ="공장명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtFrItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd frm1.txtFrItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtFrItemNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR >
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" WIDTH=100%>
									<script language =javascript src='./js/c3608ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>	
						</TABLE> 
   					</TD>

					<TD WIDTH=60% HEIGHT=100%>
						<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
						    <TR>
								<TD HEIGHT=10 WIDTH=100%>			
									<FIELDSET CLASS="CLSFLD">        							        
										<TABLE CLASS="BasicTB" CELLSPACING=0>

											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>표준원가</TD>
												<TD CLASS=TD6 NOWRAP>실제원가</TD>
												<TD CLASS=TD6 NOWRAP>원가차이</TD>
											</TR>
										</TABLE>
									</FIELDSET>      
								</TD>  
						     </TR>
						     <TR>	
								<TD HEIGHT=45 WIDTH=100%>			
									<FIELDSET CLASS="CLSFLD">        							        
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<LEGEND>총생산량 기준</LEGEND>

											
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20>생산량</TD>
												<TD CLASS=TD6 NOWRAP >
												<script language =javascript src='./js/c3608ma1_OBJECT1_txtTotQty.js'></script></TD>
												<TD CLASS=TD6 NOWRAP></TD>
												<TD CLASS=TD6 NOWRAP></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>

											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=20>재료비</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT2_txtStd_Mcost1.js'></script>
												</TD>

												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT3_txtReal_Mcost1.js'></script>
												</TD>

												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT4_txtDiff_Mcost1.js'></script>
												</TD>

											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
														
											<TR>
												<TD CLASS=TD5 NOWRAP>노무비</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT5_txtStd_Lcost1.js'></script>
												</TD>

												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT6_txtReal_Lcost1.js'></script>

												</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT7_txtDiff_Lcost1.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>경비</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT8_txtStd_Ecost1.js'></script>
												</TD>

												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT9_txtReal_Ecost1.js'></script>

												</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT10_txtDiff_Ecost1.js'></script>

											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>합계</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT11_txtStd_Sum1.js'></script>
												</TD>

												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT12_txtReal_Sum1.js'></script>

												</TD>
												<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/c3608ma1_OBJECT13_txtDiff_Sum1.js'></script>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
										</TABLE>	
									</FIELDSET>	    
								</TR>
								<TR>

									<TD HEIGHT=45 WIDTH=100%>
									<FIELDSET CLASS="CLSFLD">			
									        <TABLE CELLSPACING=0 CELLPADDING=0 WIDTH="100%" HEIGHT="100%">
										<LEGEND>단위생산량기준</LEGEND>
										
										
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>

										<TR>
											<TD CLASS=TD5 NOWRAP>재료비</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT14_txtStd_Mcost2.js'></script>
											</TD>

											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT15_txtReal_Mcost2.js'></script>

											</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT16_txtDiff_Mcost2.js'></script>

											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>
													
										<TR>
											<TD CLASS=TD5 NOWRAP>노무비</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT17_txtStd_Lcost2.js'></script>
											</TD>

											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT18_txtReal_Lcost2.js'></script>

											</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT19_txtDiff_Lcost2.js'></script>

											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>

										<TR>
											<TD CLASS=TD5 NOWRAP>경비</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT20_txtStd_Ecost2.js'></script>
											</TD>

											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT21_txtReal_Ecost2.js'></script>

											</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT22_txtDiff_Ecost2.js'></script>

										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>

										<TR>
											<TD CLASS=TD5 NOWRAP>합계</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT23_txtStd_Sum2.js'></script>
											</TD>

											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT24_txtReal_Sum2.js'></script>

											</TD>
											<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/c3608ma1_OBJECT25_txtDiff_Sum2.js'></script>

										</TR>
										<TR>
											<TD HEIGHT=20 CLASS=TD5 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>
											
											</TD>

											<TD CLASS=TD6 NOWRAP>
											</TD>
											<TD CLASS=TD6 NOWRAP>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP HEIGHT=10>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>
						    	</TABLE>	
									</FIELDSET>
									</TD>
								</TR>
						</TABLE>
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
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hFrItemCd" tag="24" TABINDEX="-1" >
<script language =javascript src='./js/c3608ma1_I737242275_txtValidFromDt.js'></script>
&nbsp;~&nbsp;
<script language =javascript src='./js/c3608ma1_I660234928_txtValidToDt.js'></script>
<INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg1" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoDefaultFlg1">
<INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg2" CLASS="RADIO" tag="24X" Value="Y" CHECKED><LABEL FOR="rdoDefaultFlg1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

