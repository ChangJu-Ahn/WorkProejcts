
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Cost
*  2. Function Name        : 실제원가조회 
*  3. Program ID           : C3603MA1
*  4. Program Name         : C3603MA1
*  5. Program Desc         : 실제원가조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/06/13
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

Const BIZ_PGM_ID 		= "C3603MB1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 

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
	Dim BaseDate
                                                 
	BaseDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date

	frm1.txtYYYYMM.text	= UniConvDateAToB(BaseDate,Parent.gServerDateFormat,Parent.gDateFormat)
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
    
	Call SetZAdoSpreadSheet("C3603MA101","G","A","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock 

End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
'    With frm1
    
'    .vspdData.ReDraw = False
'	ggoSpread.SpreadLock 1 , -1
'    .vspdData.ReDraw = True

'    End With
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

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

'    lgMaxFieldCount =  UBound(Parent.gFieldNM)                      

'    ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)

'    Call Parent.MakePopData(Parent.gDefaultT,Parent.gFieldNM,Parent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,Parent.C_MaxSelList)

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    frm1.txtYYYYMM.focus
    frm1.txtSum.allownull = False
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
	Dim strYear
	Dim strMonth
	Dim strDay

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    With frm1

        
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> Parent.OPMD_CMODE Then   ' This means that it is first search
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								'☜:조회표시 
			strVal = strVal & "&txtYYYYMM=" & .hYyyymm.value					 			    '☆: 조회 조건 데이타 
			strVal = strVal & "&txtWorkStepCd=" & Trim(.hWorkStepCd.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)						
			strVal = strVal & "&txtItemAcctCd=" & Trim(.hItemAcctCd.value)						
			strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)						
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								'☜:조회표시 
			strVal = strVal & "&txtYYYYmm=" & strYear & strMonth
			strVal = strVal & "&txtWorkStepCd=" & Trim(.txtWorkStepCd.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)			
			strVal = strVal & "&txtItemAcctCd=" & Trim(.txtItemAcctCd.value)						
			strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)						
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)						
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
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

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 

End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
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
	Else
		ggoSpread.Source=frm1.vspddata	
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=======================================================================================================
'	Name : OpenWorkStep()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
		Case 1
			arrParam(0) = "작업단계팝업"			'팝업 명칭 
			arrParam(1) = "B_MINOR M, B_CONFIGURATION C"						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtWorkStepCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = "M.MINOR_CD = C.MINOR_CD and M.MAJOR_CD = C.MAJOR_CD and C.SEQ_NO = 4 and C.REFERENCE = " & FilterVar("Y", "''", "S") & "  and M.MAJOR_CD = " & FilterVar("C2000", "''", "S") & " "							'Where Condition
			arrParam(5) = "작업단계"				'TextBox 명칭 
	
			arrField(0) = "M.minor_CD"					'Field명(0)
			arrField(1) = "M.minor_NM"					'Field명(1)
    
			arrHeader(0) = "작업단계코드"				'Header명(0)
			arrHeader(1) = "작업단계명"				'Header명(1)

		Case 2
			arrParam(0) = "공장팝업"			'팝업 명칭 
			arrParam(1) = "B_PLANT"						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPlantCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = ""							'Where Condition
			arrParam(5) = "공장"				'TextBox 명칭 
	
			arrField(0) = "PLANT_CD"					'Field명(0)
			arrField(1) = "PLANT_NM"					'Field명(1)
    
			arrHeader(0) = "공장코드"				'Header명(0)
			arrHeader(1) = "공장명"				'Header명(1)

		Case 3
			arrParam(0) = "품목계정팝업"			'팝업 명칭 
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S") & "  and A.MINOR_CD = B.ITEM_ACCT AND B.ITEM_ACCT_GROUP in (" & FilterVar("1FINAL","''","S") & "," & FilterVar("2SEMI","''","S") & ")"
			arrParam(5) = "품목계정"				'TextBox 명칭 
	
			arrField(0) = "MINOR_CD"					'Field명(0)
			arrField(1) = "MINOR_NM"					'Field명(1)
    
			arrHeader(0) = "품목계정"				'Header명(0)
			arrHeader(1) = "품목계정명"				'Header명(1)

		Case 4
			arrParam(0) = "품목그룹팝업"			'팝업 명칭 
			arrParam(1) = "B_ITEM_GROUP"						'TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemGroupCd.Value)	'Code Condition
			arrParam(3) = ""							'Name Cindition
			arrParam(4) = "LEAF_FLG =" & FilterVar("Y", "''", "S") & "  and DEL_FLG =" & FilterVar("N", "''", "S") & " "							'Where Condition
			arrParam(5) = "품목그룹"				'TextBox 명칭 
	
			arrField(0) = "ITEM_GROUP_CD"					'Field명(0)
			arrField(1) = "ITEM_GROUP_NM"					'Field명(1)
    
			arrHeader(0) = "품목그룹"				'Header명(0)
			arrHeader(1) = "품목그룹명"				'Header명(1)



	End Select
    

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet,iWhere)
	End If
		
End Function


'=======================================================================================================
'	Name : OpenItem()
'	Description : Condition Item PopUp
'=======================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	If Not chkField(Document, "1") Then
	   Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("220705","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If


	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItem(arrRet)
	End If
		
End Function

Function SetPopup(byval arrRet,byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtWorkStepCd.Value = arrRet(0)		
				.txtWorkStepNm.Value = arrRet(1)		
			Case 2
				.txtPlantCd.Value = arrRet(0)		
				.txtPlantNm.Value = arrRet(1)		
			Case 3
				.txtItemAcctCd.Value = arrRet(0)		
				.txtItemAcctNm.Value = arrRet(1)		
			Case 4
				.txtItemGroupCd.Value = arrRet(0)		
				.txtItemGroupNm.Value = arrRet(1)		
			
		End Select

		lgBlnFlgChgValue = True
	End With
End Function




'=======================================================================================================
'	Name : SetItem()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetItem(byval arrRet)
	frm1.txtItemCd.Value = arrRet(0)		
	frm1.txtItemNm.Value = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function
 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'======================================================================================================
'   Event Name : txtYYYYMM_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtYYYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtYYYYMM.Action = 7
	End If
End Sub


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
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
'	Dim ii
	
    Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================


Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
	
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실제원가조회</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/c3603ma1_txtYYYYMM_txtYYYYMM.js'></script>
									</TD>
									<TD CLASS="TD5">작업단계</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtWorkStepCd" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="작업단계"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWorkStepCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
										 <INPUT TYPE=TEXT ID="txtWorkStepNm" NAME="txtWorkStepNm" SIZE=30 tag="14X">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2)">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemAcctCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(3)">
										 <INPUT TYPE=TEXT ID="txtItemAcctNm" NAME="txtAcctNm" SIZE=30 tag="14X">
									</TD>
								</TR>               
								<TR>
									<TD CLASS="TD5">품목그룹</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemGroupCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(4)">
										 <INPUT TYPE=TEXT ID="txtItemGroupNm" NAME="txtItemGroupNm" SIZE=30 tag="14X">
									</TD>
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItem()">
										 <INPUT TYPE=TEXT ID="txtItemNm" NAME="txtItemNm" SIZE=30 tag="14X">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c3603ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>가공비합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3603ma1_fpDoubleSingle2_txtMFCSum.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>재료비합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3603ma1_fpDoubleSingle2_txtMATSum.js'></script>&nbsp;
									</TD>
								</TR>			
								<TR>
									<TD CLASS="TD5" NOWRAP>(반)제품투입합계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3603ma1_fpDoubleSingle2_txtSEMISum.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>총 계</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3603ma1_fpDoubleSingle2_txtSum.js'></script>&nbsp;
									</TD>
								</TR>			
							</TABLE>
						</FIELDSET>
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
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hYyyymm" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hWorkStepCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

