<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1213MA1
'*  4. Program Name         : 조정형(공정 외)검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "q1213mb1.asp"				'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "q1213mb2.asp"				'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_LOOKUP_ID = "q1213mb3.asp"

Const BIZ_PGM_JUMP_ID = "q1211ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_BpCd					'= 1									'☆: Spread Sheet의 Column별 상수 
Dim C_BpPopup				'= 2
Dim C_BpNm					'= 3 
Dim C_SwitchNm				'= 4
Dim C_InspLevel				'= 5
Dim C_InspLevelPopup		'= 6
Dim C_AQL					'= 7
Dim C_AQLPopup				'= 8
Dim C_SubstituteForSigmaNm	'= 9
Dim C_MthdOfDecisionNm		'= 10
'------------------ Hidden Column ------------------
Dim C_SwitchCd				'= 11
Dim C_SubstituteForSigmaCd	'= 12
Dim C_MthdOfDecisionCd		'= 13

Dim IsOpenPop						' Popup

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                            '⊙: initializes sort direction
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
	End If	
	
	If ReadCookie("txtItemNm") <> "" Then
		frm1.txtItemNm.Value = ReadCookie("txtItemNm")
	End If	

	If ReadCookie("txtInspClassCd") <> "" Then
		frm1.cboInspClassCd.Value = ReadCookie("txtInspClassCd")
	else
		frm1.cboInspClassCd.Value = "R"
	End If	
				
	If ReadCookie("txtInspItemCd") <> "" Then
		frm1.txtInspItemCd.Value = ReadCookie("txtInspItemCd")
	End If	
		
	If ReadCookie("txtInspItemNm") <> "" Then
		frm1.txtInspItemNm.Value = ReadCookie("txtInspItemNm")
	End If	
	
	If ReadCookie("txtInspMthdCd") <> "" Then
		frm1.txtInspMthdCd.Value = ReadCookie("txtInspMthdCd")
	End If	
		
	If ReadCookie("txtInspMthdNm") <> "" Then
		frm1.txtInspMthdNm.Value = ReadCookie("txtInspMthdNm")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtInspClassCd", ""
	WriteCookie "txtInspItemCd", ""
	WriteCookie "txtInspItemNm", ""
	WriteCookie "txtInspMthdCd", ""
	WriteCookie "txtInspMthdNm", ""
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	 
	 Call InitSpreadPosVariables()
	     
     With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20040518", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
    	.MaxCols = C_MthdOfDecisionCd + 1				'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		Call AppendNumberPlace("6", "4","3")
		
    	ggoSpread.SSSetEdit C_BpCd, "공급처코드", 12, 0, -1, 10, 2
    	ggoSpread.SSSetButton C_BpPopup
    	ggoSpread.SSSetEdit C_BpNm, "공급처명", 20, 0, -1, 40
    	ggoSpread.SSSetCombo C_SwitchNm, "엄격도", 10, 0, False
    	ggoSpread.SSSetEdit C_InspLevel, "검사수준", 14, 0, -1, 3, 2
    	ggoSpread.SSSetButton C_InspLevelPopup
    	ggoSpread.SSSetFloat C_AQL, "AQL", 10, 6, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    	ggoSpread.SSSetButton C_AQLPopup
    	ggoSpread.SSSetCombo C_SubstituteForSigmaNm, "표준편차대용", 15, 0, False
    	ggoSpread.SSSetCombo C_SwitchCd, "엄격도", 10, 0, False
    	ggoSpread.SSSetCombo C_MthdOfDecisionNm, "판정방법", 15, 0, False
    	ggoSpread.SSSetCombo C_SubstituteForSigmaCd, "표준편차대용", 10, 0, False
    	ggoSpread.SSSetCombo C_MthdOfDecisionCd, "판정방법", 10, 0, False
    	
    	Call ggoSpread.MakePairsColumn(C_BpCd,C_BpPopup)
    	Call ggoSpread.MakePairsColumn(C_InspLevel,C_InspLevelPopup)
    	Call ggoSpread.MakePairsColumn(C_AQL,C_AQLPopup)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
 		Call ggoSpread.SSSetColHidden(C_SwitchCd, C_SwitchCd, True)
 		Call ggoSpread.SSSetColHidden(C_SubstituteForSigmaCd, C_SubstituteForSigmaCd, True)
 		Call ggoSpread.SSSetColHidden(C_MthdOfDecisionCd, C_MthdOfDecisionCd, True)
	
	End With	
	
    frm1.vspdData.ReDraw = true
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With ggoSpread
		Call .SSSetRequired(C_BpCd, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_SwitchNm, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_InspLevel, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_AQL, pvStartRow, pvEndRow)
    	
    	Call .SSSetProtected(C_BpNm, pvStartRow, pvEndRow)    	
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " AND MINOR_CD <> " & FilterVar("P", "''", "S") & "  ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	
	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))

	With frm1.vspdData

		Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("Q0007", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
		strCboCd = ""
		strCboNm = ""
					
		strCboCd = lgF0 
		strCboNm = lgF1
				
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
					
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_SwitchCd
		ggoSpread.SetCombo strCboNm, C_SwitchNm
				
				
		Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("Q0015", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				
		strCboCd = ""
		strCboNm = ""
				
		strCboCd = vbTab & lgF0 
		strCboNm = vbTab & lgF1
				
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
					
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_SubstituteForSigmaCd
		ggoSpread.SetCombo strCboNm, C_SubstituteForSigmaNm
			
				
		Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("Q0016", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				
		strCboCd = ""
		strCboNm = ""
				
		strCboCd = vbTab & lgF0 
		strCboNm = vbTab & lgF1
				
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
					
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_MthdOfDecisionCd
		ggoSpread.SetCombo strCboNm, C_MthdOfDecisionNm	
	
	End With

End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_BpCd					= 1									'☆: Spread Sheet의 Column별 상수 
	C_BpPopup				= 2
	C_BpNm					= 3 
	C_SwitchNm				= 4
	C_InspLevel				= 5
	C_InspLevelPopup		= 6
	C_AQL					= 7
	C_AQLPopup				= 8
	C_SubstituteForSigmaNm	= 9
	C_MthdOfDecisionNm		= 10
	'------------------ Hidden Column ------------------
	C_SwitchCd				= 11
	C_SubstituteForSigmaCd	= 12
	C_MthdOfDecisionCd		= 13	

End Sub
 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_BpCd					= iCurColumnPos(1)									'☆: Spread Sheet의 Column별 상수 
		C_BpPopup				= iCurColumnPos(2)
		C_BpNm					= iCurColumnPos(3)
		C_SwitchNm				= iCurColumnPos(4)
		C_InspLevel				= iCurColumnPos(5)
		C_InspLevelPopup		= iCurColumnPos(6)
		C_AQL					= iCurColumnPos(7)
		C_AQLPopup				= iCurColumnPos(8)
		C_SubstituteForSigmaNm	= iCurColumnPos(9)
		C_MthdOfDecisionNm		= iCurColumnPos(10)
		'------------------ Hidden Column ------------------
		C_SwitchCd				= iCurColumnPos(11)
		C_SubstituteForSigmaCd	= iCurColumnPos(12)
		C_MthdOfDecisionCd		= iCurColumnPos(13)
 	End Select
End Sub

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	frm1.txtItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.Focus		
	End If	

	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

 '------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

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

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

 '------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : Inspection Item By Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	OpenInspItem = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12
	Dim iCalledAspName, IntRetCD
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	'검사분류가 있는 지 체크 
	If Trim(frm1.cboInspClassCd.Value) = "" then 
		Call DisplayMsgBox("229915", "X", "X", "X") 		'검사분류정보가 필요합니다 
		frm1.cboInspClassCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	'품목코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916", "X", "X", "X") 		'품목정보가 필요합니다 
		frm1.txtItemCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	With frm1
		Param1 = Trim(.txtPlantCd.Value)
		Param2 = Trim(.txtPlantNm.Value)
		Param3 = Trim(.txtItemCd.Value)
		Param4 = Trim(.txtItemNm.Value)
		Param5 = Trim(.cboInspClassCd.Value)
		Param6 = Trim(.cboInspClassCd.Options(.cboInspClassCd.SelectedIndex).Text)
		Param7 = ""
		Param8 = ""
		Param9 = ""
		Param10 = Trim(.txtInspItemCd.value)
		Param11 = ""
		Param12 = "0300"	'조정형 
	End With
	
	iCalledAspName = AskPRAspName("q1211pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspItemCd.Value = arrRet(1)
		frm1.txtInspItemNm.Value = arrRet(2)	
		frm1.txtInspMthdCd.Value = arrRet(3)
		frm1.txtInspMthdNm.Value = arrRet(4)
		frm1.txtInspItemCd.Focus
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function

'------------------------------------------  OpenBp()  -------------------------------------------------
'	Name : OpenBp()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode)
	OpenBp = false

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER"					' TABLE 명칭 
	arrParam(2) = strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " )"			' Where Condition
	arrParam(5) = "공급처"						' 조건필드의 라벨 명칭	
	
    arrField(0) = "BP_CD"								' Field명(0)
    arrField(1) = "BP_NM"								' Field명(1)
    
    arrHeader(0) = "공급처코드"					' Header명(0)
    arrHeader(1) = "공급처명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_BpCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BpNm
			.vspdData.Text = arrRet(1)
		
			Call vspdData_Change(C_BpCd, .vspdData.ActiveRow)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_BpCd,.vspdData.ActiveRow,"M","X","X")
		End With
	End If	
	Set gActiveElement = document.activeElement
	OpenBp = true
End Function

 '------------------------------------------  OpenInspLevel()  -------------------------------------------------
'	Name : OpenInspLevel()
'	Description : InspLevel PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspLevel(Byval strCode)
	OpenInspLevel = false

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "검사수준팝업"					' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) = strCode									' Code Condition
	arrParam(3) = ""										' Name Cindition
	If frm1.txtInspMthdCd.Value = "" then
		Call DisplayMsgBox("220906", "X", "X", "X") 		'검사방식을 선택하십시오 
		IsOpenPop = False
		Exit Function
	Else
		If Mid(frm1.txtInspMthdCd.Value, 2,1) = "3" Then			'조정형 
			If Mid(frm1.txtInspMthdCd.Value, 1,1) = "1" Then		'계수형 
				arrParam(4) = "MAJOR_CD = " & FilterVar("Q0008", "''", "S") & ""				' Where Condition
			ElseIf Mid(frm1.txtInspMthdCd.Value, 1,1) = "2" Then	'계량형 
				arrParam(4) = "MAJOR_CD = " & FilterVar("Q0022", "''", "S") & ""				' Where Condition
			Else
				Call DisplayMsgBox("220905", "X", "X", "X") 		'검사방식이 조정형이 아닙니다 
				IsOpenPop = False
				Exit Function
			End If
		Else
			Call DisplayMsgBox("220905", "X", "X", "X") 		'검사방식이 조정형이 아닙니다 
			IsOpenPop = False
			Exit Function
		End If
	End If
	
	arrParam(5) = "검사수준"							' 조건필드의 라벨 명칭	
	
    arrField(0) = "ED40" & parent.gcolsep & "MINOR_CD"								' Field명(0)
    
    arrHeader(0) = "검사수준"						' Header명(0)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData,C_InspLevel,frm1.vspdData.ActiveRow,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_InspLevel
			.vspdData.Text = arrRet(0)
			
			Call vspdData_Change(C_InspLevel, .vspdData.ActiveRow)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_InspLevel,.vspdData.ActiveRow,"M","X","X")
		End With
	End If	
	OpenInspLevel = true
End Function

'------------------------------------------  OpenAQL()  -------------------------------------------------
'	Name : OpenAQL()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAQL(Byval strCode)
	OpenAQL = false
	
	Dim arrRet
	Dim arrParam1, arrParam2
	Dim iCalledAspName, IntRetCD

	If frm1.txtInspMthdCd.Value = "" then
		Call DisplayMsgBox("220906", "X", "X", "X") 		'검사방식을 선택하십시오 
		IsOpenPop = False
		Exit Function
	Else
		If Mid(frm1.txtInspMthdCd.Value, 2,1) = "3" Then			'조정형 
			If Mid(frm1.txtInspMthdCd.Value, 1,1) = "1" Then	'계수형 
				arrParam2 = "Q0011"				' Where Condition		
			ElseIf Mid(frm1.txtInspMthdCd.Value, 1,1) = "2" Then	'계량형 
				arrParam2 = "Q0012"				' Where Condition
			Else
				Call DisplayMsgBox("220905", "X", "X", "X") 		'검사방식이 조정형이 아닙니다 
				IsOpenPop = False
				Exit Function
			End If
		Else
			Call DisplayMsgBox("220905", "X", "X", "X") 		'검사방식이 조정형이 아닙니다 
			IsOpenPop = False
			Exit Function
		End If
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = strCode
	
	iCalledAspName = AskPRAspName("q1211pa3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2), _
	              "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData,C_AQL,frm1.vspdData.ActiveRow,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_AQL
			.vspdData.Text = arrRet(0)
			
			Call vspdData_Change(C_AQL, .vspdData.ActiveRow)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_AQL,.vspdData.ActiveRow,"M","X","X")
		End With
	End If	
	OpenAQL = true
End Function

'=============================================  2.5.2 LoadInspStand()  ======================================
'=	Event Name : LoadInspStand
'=	Event Desc :
'========================================================================================================
Function LoadInspStand()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
	End With
	PgmJump(BIZ_PGM_JUMP_ID)

End Function

'================================== 2.6.1 SetSpreadColorForVariable() ====================================
' Function Name : SetSpreadColorForVariable
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadColorForVariable(ByVal pvStartRow, ByVal pvEndRow)
    With ggoSpread
		Call .SpreadUnLock(C_SubstituteForSigmaCd, pvStartRow, C_SubstituteForSigmaCd, pvEndRow)
		Call .SpreadUnLock(C_SubstituteForSigmaNm, pvStartRow, C_SubstituteForSigmaNm, pvEndRow)
		Call .SpreadUnLock(C_MthdOfDecisionCd, pvStartRow, C_MthdOfDecisionCd, pvEndRow)
		Call .SpreadUnLock(C_MthdOfDecisionNm, pvStartRow, C_MthdOfDecisionNm, pvEndRow)
		
    	Call .SSSetRequired(C_SubstituteForSigmaCd, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_SubstituteForSigmaNm, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_MthdOfDecisionCd, pvStartRow, pvEndRow)
    	Call .SSSetRequired(C_MthdOfDecisionNm, pvStartRow, pvEndRow)
    End With
End Sub

'================================== 2.6.2 SetSpreadColorForAttribute() ====================================
' Function Name : SetSpreadColorForAttribute
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadColorForAttribute(ByVal pvStartRow, ByVal pvEndRow)
	With ggoSpread
		Call .SSSetProtected(C_SubstituteForSigmaCd, pvStartRow, pvEndRow)
    	Call .SSSetProtected(C_SubstituteForSigmaNm, pvStartRow, pvEndRow)
    	Call .SSSetProtected(C_MthdOfDecisionCd, pvStartRow, pvEndRow)
    	Call .SSSetProtected(C_MthdOfDecisionNm, pvStartRow, pvEndRow)
    	
    End With
End Sub

'================================== 2.6.3 ProtectBpCd() ====================================
' Function Name : ProtectBpCd
' Function Desc : 
'=========================================================================================================
Sub ProtectBpCd(ByVal pvInspClassCd)
	Dim lRow
	
    SELECT CASE pvInspClassCd
    	CASE "R"
    		For lRow = 1 To frm1.vspdData.MaxRows
    			frm1.vspdData.Row = lRow
				frm1.vspdData.Col = 0	
					
				If frm1.vspdData.Text = ggoSpread.InsertFlag Then
					ggoSpread.SpreadUnLock C_BpCd, lRow, C_BpCd, lRow
					ggoSpread.SSSetRequired C_BpCd, lRow, lRow
					ggoSpread.SpreadUnLock C_BpPopup, lRow, C_BpPopup, lRow
				End If
			Next	
    	CASE "F"
    		ggoSpread.SSSetProtected C_BpCd, -1, -1
			ggoSpread.SSSetProtected C_BpPopup, -1, -1
    	CASE "S"
    		ggoSpread.SSSetProtected C_BpCd, -1, -1
			ggoSpread.SSSetProtected C_BpPopup, -1, -1
    	CASE ELSE
    		ggoSpread.SSSetProtected C_BpCd, -1, -1
			ggoSpread.SSSetProtected C_BpPopup, -1, -1
    END SELECT
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call InitVariables                                                      				'⊙: Initializes local global variables
	Call InitSpreadSheet                                                    			'⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolBar("11101101001011")							'⊙: 버튼 툴바 제어 
	
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    	ggoSpread.Source = frm1.vspdData
    	ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
		Select Case Col
			Case  C_SwitchNm
				.Col = Col
				intIndex = .Value
				.Col = C_SwitchCd
				.Value = intIndex
			Case  C_SubstituteForSigmaNm
				.Col = Col
				intIndex = .Value
				.Col = C_SubstituteForSigmaCd
				.Value = intIndex
			Case  C_MthdOfDecisionNm
				.Col = Col
				intIndex = .Value
				.Col = C_MthdOfDecisionCd
				.Value = intIndex
		End Select
	End With

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	 '----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_BpPopUp Then
			.Col = C_BpCd
			.Row = Row
			.Action = 0
			Call OpenBp(.Text)
		ElseIf Row > 0 And Col = C_InspLevelPopUp Then
			.Col = C_InspLevel
			.Row = Row
			.Action = 0
			Call OpenInspLevel(.Text)    
		ElseIf Row > 0 And Col = C_AQLPopup Then
			.Col = C_AQL
			.Row = Row
			.Action = 0
			Call OpenAQL(.Text)    
		End If
		
		Call SetFocusToDocument("M")
		.Focus
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If   
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
 	
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
    
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If

End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii )

End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	frm1.vspdData.Redraw = False
    
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    
    Call InitSpreadSheet
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData
	
	Call SetSpreadColor(-1, -1)
	Call ggoSpread.SSSetProtected(C_BpCd, -1)
	Call ggoSpread.SSSetProtected(C_BpPopup, -1)
	
	If Left(frm1.txtInspMthdCd.value,1) = "2" Then
		Call SetSpreadColorForVariable(1, frm1.vspdData.MaxRows)
	Else
		Call SetSpreadColorForAttribute(1, frm1.vspdData.MaxRows)
	End If
	
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	
	frm1.vspdData.Redraw = True
End Sub 

'/* 2003-07 정기패치 : 검사방식 LOOK UP 기능 추가 - START */
'=======================================================================================================
'   Event Name : txtPlantCd_OnChange
'   Event Desc : 
'=======================================================================================================
Sub txtPlantCd_OnChange()
	Dim strPlantCd
	Dim strInspClassCd
	Dim strItemCd
	Dim strInspItemCd
	
	If gLookUpEnable = False Then Exit Sub
	
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
	
	With frm1
		strPlantCd = Trim(.txtPlantCd.value)
		strInspClassCd = Trim(.cboInspClassCd.value)
		strItemCd = Trim(.txtItemCd.value)
		strInspItemCd = Trim(.txtInspItemCd.value)
    
		.txtInspMthdCd.value = ""
		.txtInspMthdNm.value = ""
		
		If strPlantCd = "" Or strInspClassCd = "" Or strItemCd = "" Or strInspItemCd = "" Then Exit Sub
    End With

    Call LayerShowHide(1)
    Call window.setTimeout("LookUpInspMethod """ + strPlantCd + """, """ + strInspClassCd + """, """ + strItemCd + """, """ + strInspItemCd + """", 1)   
End Sub

'==========================================================================================
'   Event Name :cboInspClassCd_onChange
'   Event Desc :
'==========================================================================================
Sub cboInspClassCd_onChange()
	Dim lRow
	Dim strPlantCd
	Dim strInspClassCd
	Dim strItemCd
	Dim strInspItemCd
	
	'검사분류에 따른 공급처 코드 Enalbe/Disable 처리 
	Call ProtectBpCd(frm1.cboInspClassCd.Value)
	
	If gLookUpEnable = False Then Exit Sub
	
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
	
	With frm1
		strPlantCd = Trim(.txtPlantCd.value)
		strInspClassCd = Trim(.cboInspClassCd.value)
		strItemCd = Trim(.txtItemCd.value)
		strInspItemCd = Trim(.txtInspItemCd.value)
    	
		.txtInspMthdCd.value = ""
		.txtInspMthdNm.value = ""
		
		If strPlantCd = "" Or strInspClassCd = "" Or strItemCd = "" Or strInspItemCd = "" Then Exit Sub
    End With

    Call LayerShowHide(1)
    Call window.setTimeout("LookUpInspMethod """ + strPlantCd + """, """ + strInspClassCd + """, """ + strItemCd + """, """ + strInspItemCd + """", 1)    
End Sub

'=======================================================================================================
'   Event Name : txtItemCd_OnChange
'   Event Desc : 
'=======================================================================================================
Sub txtItemCd_OnChange()
	Dim strPlantCd
	Dim strInspClassCd
	Dim strItemCd
	Dim strInspItemCd
	
	If gLookUpEnable = False Then Exit Sub
	
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
	
	With frm1
		strPlantCd = Trim(.txtPlantCd.value)
		strInspClassCd = Trim(.cboInspClassCd.value)
		strItemCd = Trim(.txtItemCd.value)
		strInspItemCd = Trim(.txtInspItemCd.value)
    
		.txtInspMthdCd.value = ""
		.txtInspMthdNm.value = ""
		
		If strPlantCd = "" Or strInspClassCd = "" Or strItemCd = "" Or strInspItemCd = "" Then Exit Sub
    End With

    Call LayerShowHide(1)
    Call window.setTimeout("LookUpInspMethod """ + strPlantCd + """, """ + strInspClassCd + """, """ + strItemCd + """, """ + strInspItemCd + """", 1)
End Sub

'=======================================================================================================
'   Event Name : txtInspItemCd_OnChange
'   Event Desc : 
'=======================================================================================================
Sub txtInspItemCd_OnChange()
	Dim strPlantCd
	Dim strInspClassCd
	Dim strItemCd
	Dim strInspItemCd
	
	If gLookUpEnable = False Then Exit Sub
	
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
	
	With frm1
		strPlantCd = Trim(.txtPlantCd.value)
		strInspClassCd = Trim(.cboInspClassCd.value)
		strItemCd = Trim(.txtItemCd.value)
		strInspItemCd = Trim(.txtInspItemCd.value)
    
		.txtInspMthdCd.value = ""
		.txtInspMthdNm.value = ""
		
		If strPlantCd = "" Or strInspClassCd = "" Or strItemCd = "" Or strInspItemCd = "" Then Exit Sub
    End With

    Call LayerShowHide(1)
    Call window.setTimeout("LookUpInspMethod """ + strPlantCd + """, """ + strInspClassCd + """, """ + strItemCd + """, """ + strInspItemCd + """", 1)
End Sub

'=======================================================================================================
'	Sub Name : LookUpInspMethod																			   
'	Sub Desc :																						
'========================================================================================================
Sub LookUpInspMethod(Byval pvPlantCd, Byval pvInspClassCd, Byval pvItemCd, Byval pvInspItemCd) 
	Call CommonQueryRs("A.INSP_METHOD_CD, B.MINOR_NM ", " Q_INSPECTION_STANDARD_BY_ITEM A, B_MINOR B ", " A.INSP_METHOD_CD = B.MINOR_CD AND A.PLANT_CD =  " & FilterVar(pvPlantCd , "''", "S") & " AND A.INSP_CLASS_CD =  " & FilterVar(pvInspClassCd , "''", "S") & " AND A.ITEM_CD =  " & FilterVar(pvItemCd , "''", "S") & " AND A.INSP_ITEM_CD =  " & FilterVar(pvInspItemCd , "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	frm1.txtInspMthdCd.value = replace(lgF0,Chr(11),"")
	frm1.txtInspMthdNm.value = replace(lgF1,Chr(11),"")
	
	Call LayerShowHide(0)
	Call LookUpInspMethodOk
End Sub

'=======================================================================================================
'	Sub Name : LookUpInspMethodOk																			   
'	Sub Desc :																						
'========================================================================================================
Sub LookUpInspMethodOk()
	If frm1.vspdData.MaxRows > 0 Then
		If Left(frm1.txtInspMthdCd.value,1) = "2" Then
			Call SetSpreadColorForVariable(1, frm1.vspdData.MaxRows)
		Else
			Call SetSpreadColorForAttribute(1, frm1.vspdData.MaxRows)
		End If	
	End If
End Sub

'/* 2003-07 정기패치 : 검사방식 LOOK UP 기능 추가 - END */

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    
    Dim IntRetCD 
    
    FncQuery = False                                                        						'⊙: Processing is NG
    
    Err.Clear                                                               						'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")  
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
    Call InitVariables										'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							'⊙: This function check indispensable field
    	Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    	
	If DbQuery = False then
		Exit Function
	End If											'☜: Query db data
       
    FncQuery = True										'⊙: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
	FncNew = False                                                          '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	'On Error Resume Next                                                    '☜: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	'-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetDefaultVal
	Call SetToolBar("11101101001011")							'⊙: 버튼 툴바 제어 
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
	FncNew = True
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    FncDelete = false
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    	
    Dim IntRetCD 
    
    FncSave = False                                                         						'⊙: Processing is NG
    
    Err.Clear                                                               						'☜: Protect system from crashing
    On Error Resume Next                                                    
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
    	IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
    	Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then   			'⊙: Check contents area
		Exit Function
	End If
	
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSDefaultCheck = false Then   			'⊙: Check contents area
		Exit Function
	End If

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then	
		Exit Function
	End If		                                                  			'☜: Save db data
    
    FncSave = True                                                          						'⊙: Processing is OK
    	
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()	
	FncCopy = false
	With frm1
		If .vspdData.MaxRows < 1 then
	    	Exit function
    	End if

		If frm1.txtInspMthdCd.Value = "" then
			Call DisplayMsgBox("220906", "X", "X", "X") 		'검사방식을 선택하십시오 
			Exit Function	
		End If

		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData	
		ggoSpread.CopyRow
		
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow
		
		If Left(.txtInspMthdCd.value,1) = "2" Then
			Call SetSpreadColorForVariable(1, .vspdData.MaxRows)
		Else
			Call SetSpreadColorForAttribute(1, .vspdData.MaxRows)
		End If
		
		.vspdData.Row = .vspdData.ActiveRow
	    .vspdData.Col = C_BpCd
	    .vspdData.Text = ""
	    .vspdData.Col = C_BpNm
	    .vspdData.Text = ""
		
		'검사분류에 따른 공급처 코드 Enalbe/Disable 처리 
		Call ProtectBpCd(.cboInspClassCd.Value)
	    	
	    .vspdData.ReDraw = True                                   					            '☜: Protect system from crashing
	End With
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste()		
	FncPaste = false
    ggoSpread.SpreadPaste
	FncPaste = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
   	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if

	ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    	
	frm1.vspdData.Focus
        	                                                  						'☜: Protect system from crashing
    FncCancel = true
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow
	Dim pvRow
	
	On Error Resume Next

	If frm1.txtInspMthdCd.Value = "" then
		Call DisplayMsgBox("220906", "X", "X", "X") 		'검사방식을 선택하십시오 
		Exit Function	
	End If
	
	FncInsertRow = false

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If

	
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
		
		If Left(.txtInspMthdCd.value,1) = "2" Then
			Call SetSpreadColorForVariable(1, .vspdData.MaxRows)
		Else
			Call SetSpreadColorForAttribute(1, .vspdData.MaxRows)
		End If
				
		'검사분류에 따른 공급처 코드 Enalbe/Disable 처리 
		Call ProtectBpCd(.cboInspClassCd.Value)
    	
    	.vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false
	Dim lDelRows
	Dim iDelRowCnt, i
    
   	With frm1
		If .vspdData.MaxRows < 1 then
			Exit function
		End if	
		.vspdData.focus
		ggoSpread.Source = .vspdData 
	    
	     '----------  Coding part  -------------------------------------------------------------   
	
		lDelRows = ggoSpread.DeleteRow
	End With
	
	FncDeleteRow = true
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
   	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	FncPrev =false
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_MULTI, False)     
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
	FncScreenSave = false
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore()		
	FncScreenRestore = false
    
'    	If ggoSpread.AllClear = True Then       		
'       		ggoSpread.LoadLayout
'    	End If

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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	DbQuery = False
	
	Err.Clear                                                               						'☜: Protect system from crashing
	Call LayerShowHide(1)
	
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001  				'☜:
			strVal = strVal & "&txtPlantCd=" & .hPlantCd.value					'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd=" & .hItemCd.value
			strVal = strVal & "&cboInspClassCd=" & .hInspClassCd.value		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtInspItemCd=" & .hInspItemCd.value			'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey					
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001   			'☜:
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)		 	'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtInspItemCd=" & Trim(.txtInspItemCd.value)	'☆: 조회 조건 데이타 
			strVal = strVal & "&cboInspClassCd=" & Trim(.cboInspClassCd.Value)		 '☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
		
		DbQuery = True                                                          				'⊙: Processing is NG
	End With
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	ggoSpread.source = frm1.vspdData
	Call SetSpreadColor(-1, -1)
	Call ggoSpread.SSSetProtected(C_BpCd, -1)
	Call ggoSpread.SSSetProtected(C_BpPopup, -1)
	
	If Left(frm1.txtInspMthdCd.value,1) = "2" Then
		Call SetSpreadColorForVariable(1, frm1.vspdData.MaxRows)
	Else
		Call SetSpreadColorForAttribute(1, frm1.vspdData.MaxRows)
	End If
		
	lgIntFlgMode = Parent.OPMD_UMODE			'⊙: Indicates that current mode is Update mode
	Call SetToolBar("11101111001111")							'⊙: 버튼 툴바 제어 
    	
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	
'	frm1.vspdData.Focus
'	Call SetActiveCell(frm1.vspdData,C_SwitchNm,frm1.vspdData.ActiveRow,"M","X","X")
'	Set gActiveElement = document.activeElement		
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt
	Dim lGrpInsCnt
	Dim lGrpDelCnt 
	Dim strDel
	Dim strVal

	Dim iLoop
	Dim iColSep
	Dim iRowSep
	Dim iMaxRows
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	Dim arrVal
	Dim arrDel

	Dim strBpCd
	Dim strSwitchCd
	Dim strInspLevelCd
	Dim strAQL
	Dim strSubstituteForSigmaCd
	Dim strMthdOfDecisionCd
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	
	iLoop       = 1 
	iColSep     = Parent.gColSep
	iRowSep     = Parent.gRowSep
	iMaxRows    = frm1.vspdData.MaxRows
	iInsertFlag = ggoSpread.InsertFlag
	iUpdateFlag = ggoSpread.UpdateFlag
	iDeleteFlag = ggoSpread.DeleteFlag                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1   
		lGrpInsCnt = 1
		lGrpDelCnt = 1 
		strVal = ""
    	strDel = ""
  
    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
				Case iInsertFlag					'☜: 신규 
					.vspdData.Col = C_BpCd
					strBpCd = Trim(.vspdData.Text)
					.vspdData.Col = C_SwitchCd
					strSwitchCd = Trim(.vspdData.Text)
					.vspdData.Col = C_InspLevel
					strInspLevelCd = Trim(.vspdData.Text)
					.vspdData.Col = C_AQL
					strAQL = UNIConvNum(Trim(.vspdData.Text), 0)
					.vspdData.Col = C_SubstituteForSigmaCd
					strSubstituteForSigmaCd = Trim(.vspdData.Text)
					.vspdData.Col = C_MthdOfDecisionCd
					strMthdOfDecisionCd = Trim(.vspdData.Text)

					strVal = strVal & "C" & iColSep & _
									strBpCd						& iColSep & _
									strSwitchCd					& iColSep & _
									strInspLevelCd				& iColSep & _
									strAQL						& iColSep & _
									strSubstituteForSigmaCd		& iColSep & _
									strMthdOfDecisionCd			& iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpInsCnt = lGrpInsCnt + 1
					ReDim Preserve arrVal(lGrpInsCnt - 1)
					arrVal(lGrpInsCnt - 1) = strVal
					
				Case iUpdateFlag					'☜: 신규 
					.vspdData.Col = C_BpCd
					strBpCd = Trim(.vspdData.Text)
					.vspdData.Col = C_SwitchCd
					strSwitchCd = Trim(.vspdData.Text)
					.vspdData.Col = C_InspLevel
					strInspLevelCd = Trim(.vspdData.Text)
					.vspdData.Col = C_AQL
					strAQL = UNIConvNum(Trim(.vspdData.Text), 0)
					.vspdData.Col = C_SubstituteForSigmaCd
					strSubstituteForSigmaCd = Trim(.vspdData.Text)
					.vspdData.Col = C_MthdOfDecisionCd
					strMthdOfDecisionCd = Trim(.vspdData.Text)

					strVal = strVal & "U" & iColSep & _
									strBpCd						& iColSep & _
									strSwitchCd					& iColSep & _
									strInspLevelCd				& iColSep & _
									strAQL						& iColSep & _
									strSubstituteForSigmaCd		& iColSep & _
									strMthdOfDecisionCd			& iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpInsCnt = lGrpInsCnt + 1
					ReDim Preserve arrVal(lGrpInsCnt - 1)
					arrVal(lGrpInsCnt - 1) = strVal

				Case iDeleteFlag					'☜: 삭제 
					.vspdData.Col = C_BpCd
					strBpCd = Trim(.vspdData.Text)
					
					strDel = strDel & "D" & iColSep & _
									strBpCd & iColSep & _
									CStr(lRow) & iRowSep
					lGrpCnt = lGrpCnt + 1
					lGrpDelCnt = lGrpDelCnt + 1
					ReDim Preserve arrDel(lGrpDelCnt - 1)
					arrDel(lGrpDelCnt - 1) = strDel	
			End Select
		Next
	
		strVal = Join(arrVal,iRowSep)
		strDel = Join(arrDel,iRowSep)
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'☜: 비지니스 ASP 를 가동 
	End With
	
	DbSave = True                                                     '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()									'☆: 저장 성공후 실행 로직 
	DbSaveOk = false
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
	DbSaveOk = false
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = false
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()									'☆: 삭제 성공후 실행 로직 
	DbDeleteOk = false
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>조정형(공정 외) 검사조건</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
        							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        							<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
        							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE="20" MAXLENGTH="20" tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
										<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE WIDTH="100%" HEIGHT=100% <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>검사방식</TD>
								<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtInspMthdCd" SIZE="10" MAXLENGTH="4" ALT="검사방식" tag="14">
								<INPUT TYPE=TEXT NAME="txtInspMthdNm" SIZE="40" MAXLENGTH="40" tag="14" ></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT="100%" COLSPAN=2>
									<script language =javascript src='./js/q1213ma1_I334936082_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</td>
    					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspStand">검사기준</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
    				</TR>
    			</TABLE>
    		</TD>
    	</TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hInspItemCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

