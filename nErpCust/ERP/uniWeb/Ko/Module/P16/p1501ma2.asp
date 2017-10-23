<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
*  1. Module Name          : Production
*  2. Function Name        : Multi Sample
*  3. Program ID           : p1501ma2
*  4. Program Name         : p1501ma2
*  5. Program Desc         : 자원조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/27
*  8. Modified date(Last)  : 2002/11/19
*  9. Modifier (First)     : Jung Yu Kyung
* 10. Modifier (Last)      : Ryu Sung Won
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->
<!--========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<Script Language="VBScript">

Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "p1501mb9.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------
Dim C_Resource
Dim C_ResourceNm
Dim C_ResourceType
Dim C_ResourceGroup
Dim C_ResourceGroupNm
Dim C_NoOfResource
Dim C_Efficiency
Dim C_Utilization
Dim C_RunRccp
Dim C_RunCrp
Dim C_OverloadTol
Dim C_RscBaseQty
Dim C_RscBaseUnit
Dim C_MfgCost
Dim C_ValidFromDt
Dim C_ValidToDt

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow

Dim iDBSYSDate
Dim StartDate, EndDate
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Function Name : initSpreadPosVariables()	
' Function Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Resource = 1
	C_ResourceNm = 2
	C_ResourceGroup = 3
	C_ResourceGroupNm = 4
	C_ResourceType = 5
	C_NoOfResource = 6
	C_Efficiency = 7
	C_Utilization = 8
	C_RunRccp = 9
	C_RunCrp = 10
	C_OverloadTol = 11
	C_RscBaseQty = 12
	C_RscBaseUnit = 13
	C_MfgCost = 14	
	C_ValidFromDt = 15
	C_ValidToDt = 16
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	 
	frm1.txtBeginEndDt.Text	= StartDate
	frm1.txtFinishEndDt.Text= UniConvDateAToB("2999-12-31", parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.rdoRunRccp0.checked = true
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream       = UCase(Trim(Frm1.txtPlantCd.Value)) & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(frm1.txtBeginEndDt.text) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtFinishEndDt.text) & parent.gColSep

    If frm1.rdoRunRccp1.checked = true Then
		lgKeyStream   = lgKeyStream & Trim(frm1.rdoRunRccp1.value) & parent.gColSep
	ElseIf frm1.rdoRunRccp2.checked = true Then
		lgKeyStream   = lgKeyStream & Trim(frm1.rdoRunRccp2.value) & parent.gColSep
	Else
		lgKeyStream   = lgKeyStream & Trim(frm1.rdoRunRccp0.value) & parent.gColSep
	End If
End Sub
	
'========================================================================================================
' Function Name : InitComboBox()
' Function Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

End Sub

'========================================================================================================
' Function Name : InitData()
' Function Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1 , 1)
		Frm1.vspdData.focus
        Set gActiveElement = document.ActiveElement
	End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	On Error Resume Next
	Err.Clear()	

	Call InitSpreadPosVariables()

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030423",,Parent.gAllowDragDropSpread

		.MaxCols   = C_ValidToDt + 1                                          ' ☜:☜: Add 1 to Maxcols
		.MaxRows = 0
		.ReDraw = false
	  
		Call GetSpreadColumnPos("A")
								'ColumnPosition			Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		ggoSpread.SSSetEdit 	C_Resource,				"자원",			15
		ggoSpread.SSSetEdit 	C_ResourceNm,			"자원명",		25
		ggoSpread.SSSetEdit 	C_ResourceGroup,		"자원그룹",		18
		ggoSpread.SSSetEdit 	C_ResourceGroupNm,		"자원그룹명",	20
		ggoSpread.SSSetEdit 	C_ResourceType,			"자원구분",		10
		ggoSpread.SSSetFloat	C_NoOfResource,			"자원수",		10,	"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_Efficiency,			"효율",			10,	"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_Utilization,			"가동율",		10,	"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_RunRccp,				"RCCP부하계산대상",	10
		ggoSpread.SSSetEdit 	C_RunCrp,				"CRP부하계산대상",	10
		ggoSpread.SSSetFloat	C_OverloadTol,			"과부하허용율",	10,	"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_RscBaseQty,			"자원기준수량",	10,	"3",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_RscBaseUnit,			"자원기준단위",	10
		ggoSpread.SSSetFloat	C_MfgCost,				"기준단위당 단위제조경비", 10, "4",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate 	C_ValidFromDt,			"시작일",		13,	2,	parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,			"종료일",		13,	2,	parent.gDateFormat

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)
		
		ggoSpread.SSSetSplit2(1)
		.ReDraw = true
		Call SetSpreadLock 
    
    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Resource			= iCurColumnPos(1)
			C_ResourceNm		= iCurColumnPos(2)
			C_ResourceGroup		= iCurColumnPos(3)
			C_ResourceGroupNm	= iCurColumnPos(4)
			C_ResourceType		= iCurColumnPos(5)
			C_NoOfResource		= iCurColumnPos(6)
			C_Efficiency		= iCurColumnPos(7)
			C_Utilization		= iCurColumnPos(8)
			C_RunRccp			= iCurColumnPos(9)
			C_RunCrp			= iCurColumnPos(10)
			C_OverloadTol		= iCurColumnPos(11)
			C_RscBaseQty		= iCurColumnPos(12)
			C_RscBaseUnit		= iCurColumnPos(13)
			C_MfgCost			= iCurColumnPos(14)
			C_ValidFromDt		= iCurColumnPos(15)
			C_ValidToDt			= iCurColumnPos(16)
	
	End Select

End Sub	

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                                  'Initializes local global variables

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFinishEndDt.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If

    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt

    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If Not(ValidDateCheck(frm1.txtBeginEndDt, frm1.txtFinishEndDt)) Then
        frm1.txtFinishEndDt.focus
        Exit Function
    End If
    
    Call MakeKeyStream("X")

    If DbQuery = False Then
        Exit Function
    End If
 
    FncQuery = True																'☜: Processing is OK
End Function
	
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
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

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
 
    FncExit = True

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    LayerShowHide(1) 
		
    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar


	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")
		
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If
	
	frm1.vspdData.Row = row
	If lgOldRow <> Row Then
		frm1.vspdData.Col = 1
		'frm1.vspdData.Row = row
		lgOldRow = Row
		  		
		With frm1
			.vspdData.Row = .vspdData.ActiveRow 	

			.vspdData.Col = C_Resource
			.txtResource2Cd.value = .vspdData.Text
		
			.vspdData.Col = C_ResourceNm
			.txtResource2Nm.value = .vspdData.Text
		
			.vspdData.Col = C_ResourceGroup
			.txtResourceGroup2Cd.value = .vspdData.Text
		
			.vspdData.Col = C_ResourceGroupNm
			.txtResourceGroup2Nm.value = .vspdData.Text

			.vspdData.Col = C_ResourceType
			.txtResourceType.value = .vspdData.Text
	
			.vspdData.Col = C_NoOfResource
			.txtNoOfResource.text = .vspdData.Text

			.vspdData.Col = C_Efficiency
			.txtEfficiency.text = .vspdData.Text

			.vspdData.Col = C_Utilization
			.txtUtilization.text = .vspdData.Text

			.vspdData.Col = C_RunRccp
			If .vspdData.Text = "Y" Then
				.rdoRunRccpR1.checked = True
			Else
				.rdoRunRccpR2.checked = True
			End If

			.vspdData.Col = C_RunCrp
			If .vspdData.Text = "Y" Then
				.rdoRunCrpR1.checked = True
			Else
				.rdoRunCrpR2.checked = True
			End If

			.vspdData.Col = C_OverloadTol
			.txtOverloadTol.text = .vspdData.Text

			.vspdData.Col = C_RscBaseQty
			.txtRscBaseQty.text = .vspdData.Text
			.txtResourceEa1.text = .vspdData.Text

			.vspdData.Col = C_RscBaseUnit
			.txtRscBaseUnit.Value = .vspdData.Text
			.txtResourceUnitCd1.Value = .vspdData.Text

			.vspdData.Col = C_MfgCost
			.txtMfgCost.text = .vspdData.Text									

			.vspdData.Col = C_ValidFromDt
			.txtValidFromDt.text = .vspdData.Text

			.vspdData.Col = C_ValidToDt
			.txtValidToDt.text = .vspdData.Text
		End With   
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
    	End If
    End if
End Sub
'=======================================================================================================
'   Event Name : txtFinishEndDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFinishEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFinishEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFinishEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
 Sub txtBeginEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBeginEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBeginEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFinishEndDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtBeginEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBeginEndDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원정보조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>종료일</TD>
									<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p1501ma2_txtBeginEndDt_txtBeginEndDt.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/p1501ma2_txtFinishEndDt_txtFinishEndDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>RCCP부하계산대상</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="1X" ID="rdoRunRccp0" VALUE="A"><LABEL FOR="rdoRunRccp0">전체</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="1X" ID="rdoRunRccp1" VALUE="Y"><LABEL FOR="rdoRunRccp1">예</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="1X" ID="rdoRunRccp2" VALUE="N"><LABEL FOR="rdoRunRccp2">아니오</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<!--<TD CLASS=TD5 NOWRAP>CRP부하계산대상</TD>
									<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="1X" ID="rdoRunCrp0" VALUE="A"><LABEL FOR="rdoRunCrp0">전체</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="1X" ID="rdoRunCrp1" VALUE="Y"><LABEL FOR="rdoRunCrp1">예</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="1X" ID="rdoRunCrp2" VALUE="N"><LABEL FOR="rdoRunCrp2">아니오</LABEL></TD>-->
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=50%>
									<script language =javascript src='./js/p1501ma2_OBJECT1_vspdData.js'></script>
								</TD>
								
								<TD HEIGHT=* WIDTH=50%>
										<TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResource2Cd" SIZE=20 MAXLENGTH=10 tag="24" ALT="자원">&nbsp;<INPUT TYPE=TEXT NAME="txtResource2Nm" SIZE=30 MAXLENGTH=40 tag="24" ALT="자원명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원그룹</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroup2Cd" SIZE=20 MAXLENGTH=10 tag="24" ALT="자원그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroup2Nm" SIZE=30 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원구분</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceType" SIZE=20 MAXLENGTH=10 tag="24" ALT="자원구분"></TD>																								
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원수</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I696635025_txtNoOfResource.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>효율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I201209509_txtEfficiency.js'></script>&nbsp;%
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>가동율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I540019017_txtUtilization.js'></script>&nbsp;%
												</TD>
											</TR>		
											<TR ID=Q1>
												<TD CLASS=TD5 NOWRAP>RCCP부하계산대상</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccpR" TAG="24" ID="rdoRunRccpR1" VALUE="Y"><LABEL FOR="rdoRunRccpR1">예</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccpR" TAG="24" ID="rdoRunRccpR2" VALUE="N"><LABEL FOR="rdoRunRccpR2">아니오</LABEL></TD>
											</TR>
											<TR ID=Q2>
												<TD CLASS=TD5 NOWRAP>CRP부하계산대상</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrpR" TAG="24" ID="rdoRunCrpR1" VALUE="Y"><LABEL FOR="rdoRunCrpR1">예</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrpR" TAG="24" ID="rdoRunCrpR2" VALUE="N"><LABEL FOR="rdoRunCrpR2">아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>과부하허용율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I659741616_txtOverloadTol.js'></script>&nbsp;%
												</TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>자원기준수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I210870450_txtRscBaseQty.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원기준단위</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRscBaseUnit" SIZE=5 MAXLENGTH=10 tag="24" ALT="자원기준단위"></TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>기준단위당 단위제조경비</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>															
																<script language =javascript src='./js/p1501ma2_I870344054_txtMfgCost.js'></script>
															</TD>
															<TD>											
																&nbsp;<INPUT TYPE=TEXT NAME="txtCurCd" tag=24 SIZE=5 MAXLENGTH=3 ALT="통화코드">&nbsp;/&nbsp;
															</TD>
															<TD>
																<script language =javascript src='./js/p1501ma2_I222092920_txtResourceEa1.js'></script>												
															</TD>
															<TD>
																&nbsp;<INPUT TYPE=TEXT NAME="txtResourceUnitCd1" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="자원기준단위">
															</TD>
														</TR>
													</TABLE>												
												</TD>
											</TR>																																			
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1501ma2_I544041687_txtValidFromDt.js'></script>
													&nbsp;~&nbsp;
													<script language =javascript src='./js/p1501ma2_I685246849_txtValidToDt.js'></script>										
												</TD>
											</TR>
										</TABLE>
									
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
