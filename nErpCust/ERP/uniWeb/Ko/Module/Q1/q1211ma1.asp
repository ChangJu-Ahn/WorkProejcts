<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211MA1
'*  4. Program Name         : 품목별 검사기준 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG120,PQBG110
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID =  "q1211mb1.asp"					'☆: 조회 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "q1211mb2.asp"					'☆: 저장 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP1_ID = "q1212ma1"
Const BIZ_PGM_JUMP2_ID = "q1213ma1"
Const BIZ_PGM_JUMP3_ID = "q1214ma1"
Const BIZ_PGM_JUMP4_ID = "q1215ma1"
Const BIZ_PGM_JUMP5_ID = "q1216ma1"						           '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspItemCd			'= 1							'☆: Spread Sheet의 Column별 상수 
Dim C_InspItemPopup			'= 2							'☆: Spread Sheet의 Column별 상수  
Dim C_InspItemNm			'= 3
Dim C_InspCharNm			'= 4
Dim C_InspOrder				'= 5
Dim C_InspMthdCd			'= 6
Dim C_InspMthdPopup			'= 7
Dim C_InspMthdNm			'= 8
Dim C_InspUnitIndctnNm		'= 9
Dim C_WeightNm				'= 10
Dim C_InspSpec				'= 11
Dim C_LSL					'= 12
Dim C_USL					'= 13
Dim C_MthdOfCLCalNm			'= 14
Dim C_CalculatedQty			'= 15
Dim C_LCL					'= 16
Dim C_UCL					'= 17
Dim C_MeasmtEquipmtCd		'= 18
Dim C_MeasmtEquipmtPopup	'= 19
Dim C_MeasmtEquipmtNm		'= 20
Dim C_MeasmtUnitCd			'= 21
Dim C_MeasmtUnitPopup		'= 22
Dim C_InspProcessDesc		'= 23
Dim C_Remark				'= 24
'------------------Hidden Column--------------------------
Dim C_InspCharCd			'= 25
Dim C_InspUnitIndctnCd		'= 26
Dim C_WeightCd				'= 27
Dim C_MthdOfCLCalCd			'= 28
'------------------------------------------

Dim lgInsertFlag
Dim lgUpdateFlag
Dim lgDeleteFlag

Dim IsOpenPop          

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

	ggoSpread.Source = frm1.vspdData	
	lgInsertFlag = ggoSpread.InsertFlag
	lgUpdateFlag = ggoSpread.UpdateFlag
	lgDeleteFlag = ggoSpread.DeleteFlag
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
	
	frm1.cboInspClassCd.value		= "R"
	
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
'	frm1.txtItemCd.value			= "10001"
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
	End If	
	
'	frm1.txtItemNm.value			= "조수산화알"
	If ReadCookie("txtItemNm") <> "" Then
		frm1.txtItemNm.Value = ReadCookie("txtItemNm")
	End If		
		
	If ReadCookie("txtInspClassCd") <> "" Then
		frm1.cboInspClassCd.Value = ReadCookie("txtInspClassCd")
	End If	
	
	If ReadCookie("txtInspClassCd") = "P" Then
		If ReadCookie("txtRoutNo") <> "" Then
			frm1.txtRoutNo.Value = ReadCookie("txtRoutNo")
		End If
		
		If ReadCookie("txtRoutNoDesc") <> "" Then
			frm1.txtRoutNoDesc.Value = ReadCookie("txtRoutNoDesc")
		End If
		
		If ReadCookie("txtOprNoDesc") <> "" Then
			frm1.txtOprNoDesc.Value = ReadCookie("txtOprNoDesc")
		End If
		
		If ReadCookie("txtOprNo") <> "" Then
			frm1.txtOprNo.Value = ReadCookie("txtOprNo")
		End If
	End If
		
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtInspClassCd", ""
	WriteCookie "txtRoutNo", ""
	WriteCookie "txtRoutNoDesc", ""
	WriteCookie "txtOprNo", ""
	WriteCookie "txtOprNoDesc", ""
	
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
		
		.MaxCols = C_MthdOfCLCalCd + 1			'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		Call AppendNumberPlace("6", "5","0")
		Call AppendNumberPlace("7", "3","0")
		Call AppendNumberPlace("8", "10","4")
		
		ggoSpread.SSSetEdit C_InspItemCd, "검사항목코드", 14, 0, -1, 5, 2		
		ggoSpread.SSSetButton C_InspItemPopup		
		ggoSpread.SSSetEdit C_InspItemNm, "검사항목명 ",20, 0, -1, 40
		ggoSpread.SSSetEdit C_InspCharCd, "표시속성코드", 10, 0, -1, 1
		ggoSpread.SSSetEdit C_InspCharNm, "표시속성", 10, 0, -1, 40 		
		ggoSpread.SSSetFloat C_InspOrder, "검사순서", 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "P"
		ggoSpread.SSSetEdit C_InspMthdCd, "검사방식코드", 14, 0, -1, 4, 2
		ggoSpread.SSSetButton C_InspMthdPopup
		ggoSpread.SSSetEdit C_InspMthdNm, "검사방식명", 20, 0, -1, 40
		ggoSpread.SSSetCombo C_InspUnitIndctnCd, "검사단위 품질표시코드", 5, 0, False
		ggoSpread.SSSetCombo C_InspUnitIndctnNm, "검사단위 품질표시", 10, 0, False
		ggoSpread.SSSetCombo C_WeightCd, "중요도", 5, 0, False
		ggoSpread.SSSetCombo C_WeightNm, "중요도", 10, 0, False
		ggoSpread.SSSetEdit C_InspSpec , "검사규격", 20, 2, -1, 40
		ggoSpread.SSSetFloat C_LSL, "하한규격", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_USL, "상한규격", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetCombo C_MthdOfCLCalCd, "관리한계산출방법코드", 5, 0, False
		ggoSpread.SSSetCombo C_MthdOfCLCalNm, "관리한계산출방법", 18, 0, False
		ggoSpread.SSSetFloat C_CalculatedQty, "관리한계계산수", 16, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat C_LCL, "관리하한", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_UCL, "관리상한", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit C_MeasmtEquipmtCd, "측정기코드", 20, 0, -1, 10, 2
		ggoSpread.SSSetButton C_MeasmtEquipmtPopup
		ggoSpread.SSSetEdit C_MeasmtEquipmtNm , "측정기명", 20, 0, -1, 40
		ggoSpread.SSSetEdit C_MeasmtUnitCd, "측정단위", 14, 0, -1, 3
		ggoSpread.SSSetButton C_MeasmtUnitPopup
		ggoSpread.SSSetEdit C_InspProcessDesc , "검사방법", 60, 0, -1, 400
		ggoSpread.SSSetEdit C_Remark , "비고", 40, 0, -1, 200
		
		Call ggoSpread.MakePairsColumn(C_InspItemCd, C_InspItemPopup)
		Call ggoSpread.MakePairsColumn(C_InspMthdCd, C_InspMthdPopup)
		Call ggoSpread.MakePairsColumn(C_MeasmtEquipmtCd, C_MeasmtEquipmtPopup)
		Call ggoSpread.MakePairsColumn(C_MeasmtUnitCd, C_MeasmtUnitPopup)
		
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
 		Call ggoSpread.SSSetColHidden( C_InspCharCd, C_InspCharCd, True)
 		Call ggoSpread.SSSetColHidden( C_InspUnitIndctnCd, C_InspUnitIndctnCd, True)
 		Call ggoSpread.SSSetColHidden( C_WeightCd, C_WeightCd, True)
 		Call ggoSpread.SSSetColHidden( C_MthdOfCLCalCd, C_MthdOfCLCalCd, True)

		.ReDraw = true
		
		Call SetSpreadLock
	End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_InspItemCd, -1, C_InspItemCd
		ggoSpread.SpreadLock C_InspItemPopup, -1, C_InspItemPopup
		ggoSpread.SpreadLock C_InspItemNm, -1, C_InspItemNm
		ggoSpread.SpreadLock C_InspCharNm, -1, C_InspCharNm
		ggoSpread.SpreadLock C_InspMthdCd, -1, C_InspMthdCd
		ggoSpread.SpreadLock C_InspMthdPopup, -1, C_InspMthdPopup
		ggoSpread.SpreadLock C_InspMthdNm, -1, C_InspMthdNm
		ggoSpread.SpreadLock C_MeasmtEquipmtNm, -1, C_MeasmtEquipmtNm
		
		ggoSpread.SSSetRequired C_InspOrder, -1
		ggoSpread.SSSetRequired C_InspUnitIndctnNm, -1
		Call ggoSpread.SpreadLock(frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols)
		.vspdData.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired C_InspItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspCharNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_InspOrder, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_InspMthdCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InspMthdNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_InspUnitIndctnNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MeasmtEquipmtNm, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
	End With
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm
    
    Dim strCboCd1 
    Dim strCboNm1
    
    Dim strCboCd2 
    Dim strCboNm2

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))
	
			
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0024", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	With frm1.vspdData
			
		strCboCd = lgF0 
		strCboNm = lgF1
	
		strCboCd=replace(strCboCd,Chr(11),vbTab)
		strCboNm=replace(strCboNm,Chr(11),vbTab)
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd, C_InspUnitIndctnCd

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboNm, C_InspUnitIndctnNm
	END WITH
		
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0005", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	With frm1.vspdData
			
		strCboCd1 = lgF0 
		strCboNm1 = lgF1
	
		strCboCd1 =replace(strCboCd1,Chr(11),vbTab)
		strCboNm1 =replace(strCboNm1,Chr(11),vbTab)
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd1, C_WeightCd 

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboNm1, C_WeightNm
	END WITH
	
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0017", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	With frm1.vspdData
			
		strCboCd2 = lgF0 
		strCboNm2 = lgF1
	
		strCboCd2 =replace(strCboCd2,Chr(11),vbTab)
		strCboNm2 =replace(strCboNm2,Chr(11),vbTab)
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboCd2, C_MthdOfCLCalCd 

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo strCboNm2, C_MthdOfCLCalNm
	END WITH
		    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()

	C_InspItemCd			= 1							'☆: Spread Sheet의 Column별 상수 
	C_InspItemPopup			= 2							'☆: Spread Sheet의 Column별 상수  
	C_InspItemNm			= 3
	C_InspCharNm			= 4
	C_InspOrder				= 5
	C_InspMthdCd			= 6
	C_InspMthdPopup			= 7
	C_InspMthdNm			= 8
	C_InspUnitIndctnNm		= 9
	C_WeightNm				= 10
	C_InspSpec				= 11
	C_LSL					= 12
	C_USL					= 13
	C_MthdOfCLCalNm			= 14
	C_CalculatedQty			= 15
	C_LCL					= 16
	C_UCL					= 17
	C_MeasmtEquipmtCd		= 18
	C_MeasmtEquipmtPopup	= 19
	C_MeasmtEquipmtNm		= 20
	C_MeasmtUnitCd			= 21
	C_MeasmtUnitPopup		= 22
	C_InspProcessDesc		= 23
	C_Remark				= 24
	'------------------Hidden Column--------------------------
	C_InspCharCd			= 25
	C_InspUnitIndctnCd		= 26
	C_WeightCd				= 27
	C_MthdOfCLCalCd			= 28
	'------------------------------------------
		
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
 		
		C_InspItemCd			= iCurColumnPos(1)							
		C_InspItemPopup			= iCurColumnPos(2)							
		C_InspItemNm			= iCurColumnPos(3)
		C_InspCharNm			= iCurColumnPos(4)
		C_InspOrder				= iCurColumnPos(5)
		C_InspMthdCd			= iCurColumnPos(6)
		C_InspMthdPopup			= iCurColumnPos(7)
		C_InspMthdNm			= iCurColumnPos(8)
		C_InspUnitIndctnNm		= iCurColumnPos(9)
		C_WeightNm				= iCurColumnPos(10)
		C_InspSpec				= iCurColumnPos(11)
		C_LSL					= iCurColumnPos(12)
		C_USL					= iCurColumnPos(13)
		C_MthdOfCLCalNm			= iCurColumnPos(14)
		C_CalculatedQty			= iCurColumnPos(15)
		C_LCL					= iCurColumnPos(16)
		C_UCL					= iCurColumnPos(17)
		C_MeasmtEquipmtCd		= iCurColumnPos(18)
		C_MeasmtEquipmtPopup	= iCurColumnPos(19)
		C_MeasmtEquipmtNm		= iCurColumnPos(20)
		C_MeasmtUnitCd			= iCurColumnPos(21)
		C_MeasmtUnitPopup		= iCurColumnPos(22)
		C_InspProcessDesc		= iCurColumnPos(23)
		C_Remark				= iCurColumnPos(24)
		'------------------Hidden Column--------------------------
		C_InspCharCd			= iCurColumnPos(25)
		C_InspUnitIndctnCd		= iCurColumnPos(26)
		C_WeightCd				= iCurColumnPos(27)
		C_MthdOfCLCalCd			= iCurColumnPos(28)
		'------------------------------------------

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
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	

	frm1.txtItemCd.Focus
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
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

 '------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : InspItemPlant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem(Byval strCode)
	OpenInspItem = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "검사항목 팝업"		' 팝업 명칭 
	arrParam(1) = "Q_INSPECTION_ITEM, B_MINOR"		' TABLE 명칭 
	arrParam(2) = strCode			' Code Condition
	arrParam(3) = ""				' Name Cindition
	arrParam(4) = "Q_INSPECTION_ITEM.INSP_CHAR=B_MINOR.MINOR_CD"				' Where Condition
	arrParam(4) = arrParam(4) & " AND B_MINOR.MAJOR_CD=" & FilterVar("Q0023", "''", "S") & ""				
	arrParam(5) = "검사항목"			
	
	arrField(0) = "INSP_ITEM_CD"		
	arrField(1) = "INSP_ITEM_NM"		
	arrField(2) = "INSP_CHAR"		
	arrField(3) = "MINOR_NM"		
	
	arrHeader(0) = "검사항목코드"	
	arrHeader(1) = "검사항목명"	
    arrHeader(2) = "품질특성코드"
    arrHeader(3) = "품질특성명"
         
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
	
	IsOpenPop = False
	
	With frm1
		Call SetActiveCell(.vspdData,C_InspItemCd,.vspdData.ActiveRow,"M","X","X")
		If arrRet(0) <> "" Then
			.vspdData.Col = C_InspItemCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_InspItemNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_InspCharCd
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_InspCharNm
			.vspdData.Text = arrRet(3)
				
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_InspItemCd,.vspdData.ActiveRow,"M","X","X")
		End If	
	End With
	
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function

 '------------------------------------------  OpenInspMthd()  -------------------------------------------------
'	Name : OpenInspMthd()
'	Description : Insp Method PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspMthd(Byval strCode)
	
	OpenInspMthd = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "검사방식 팝업"		' 팝업 명칭 
	arrParam(1) = "B_MINOR"			' TABLE 명칭 
	arrParam(2) = strCode			' Code Condition
	arrParam(3) = ""				' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("Q0004", "''", "S") & ""		' Where Condition
	frm1.vspdData.Col = C_InspCharCd
	If Trim(frm1.vspdData.Text) = "C" Then		'품질특성이 정량인 경우, 계량형이 아닌 모든 검사방식 
		arrParam(4) = arrParam(4) & " AND MINOR_CD LIKE " & FilterVar("[^2]%", "''", "S") & ""		'Patch(9월3주차): %5B%5E2]%25 --> [^2]%로 변경 
	End If
	arrParam(5) = "검사방식"			
	
	arrField(0) = "MINOR_CD"			' Field명(0)
	arrField(1) = "MINOR_NM"			' Field명(1)
    
	arrHeader(0) = "검사방식코드"			' Header명(0)
	arrHeader(1) = "검사방식명"		' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		Call SetActiveCell(.vspdData,C_InspMthdCd,.vspdData.ActiveRow,"M","X","X")
		If arrRet(0) <> "" Then
			.vspdData.Col = C_InspMthdCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_InspMthdNm
			.vspdData.Text = arrRet(1)
			
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_InspMthdCd,.vspdData.ActiveRow,"M","X","X")		
		End If	
	End With
	OpenInspMthd = true
End Function

 '------------------------------------------  OpenMeasmtEquipmt()  -------------------------------------------------
'	Name : OpenMeasmtEquipmt()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMeasmtEquipmt(Byval strCode)
	OpenMeasmtEquipmt = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "측정기 팝업"		' 팝업 명칭 
	arrParam(1) = "Q_MEASUREMENT_EQUIPMENT"	' TABLE 명칭 
	arrParam(2) = strCode			' Code Condition
	arrParam(3) = ""				' Name Cindition
	arrParam(4) = ""				' Where Condition
	arrParam(5) = "측정기기"			
	
	arrField(0) = "MEASMT_EQUIPMT_CD"		' Field명(0)
	arrField(1) = "MEASMT_EQUIPMT_NM"		' Field명(1)
    
	arrHeader(0) = "측정기기코드"			' Header명(0)
	arrHeader(1) = "측정기기명"		' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		Call SetActiveCell(.vspdData,C_MeasmtEquipmtCd,.vspdData.ActiveRow,"M","X","X")	
		If arrRet(0) <> "" Then
			.vspdData.Col = C_MeasmtEquipmtCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_MeasmtEquipmtNm
			.vspdData.Text = arrRet(1)
			
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_MeasmtEquipmtCd,.vspdData.ActiveRow,"M","X","X")		
		End If	
	End With
	OpenMeasmtEquipmt = true
End Function

'------------------------------------------  OpenMeasmtUnit()  -------------------------------------------------
'	Name : OpenMeasmtUnit()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMeasmtUnit(Byval strCode)
	OpenMeasmtUnit = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "단위"
	
	arrField(0) = "UNIT"	
	arrField(1) = "UNIT_NM"	
	
	arrHeader(0) = "단위코드"		
	arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		Call SetActiveCell(.vspdData,C_MeasmtUnitCd,.vspdData.ActiveRow,"M","X","X")		
		If arrRet(0) <> "" Then
			.vspdData.Col = C_MeasmtUnitCd
			.vspdData.Text = arrRet(0)
			
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)		 ' 변경이 읽어났다고 알려줌 
			Call SetActiveCell(.vspdData,C_MeasmtUnitCd,.vspdData.ActiveRow,"M","X","X")		
		End If	
	End With
	
	OpenMeasmtUnit = true
End Function

'------------------------------------------  OpenRoutNo()  -------------------------------------------------
'	Name : OpenRoutNo()
'	Description : RoutNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") 	
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ROUT_NO"							' Field명(0)
    arrField(1) = "DESCRIPTION"						' Field명(1)
    arrField(2) = "BOM_NO"							' Field명(1)
    arrField(3) = "MAJOR_FLG"						' Field명(1)
   
    arrHeader(0) = "라우팅"						' Header명(0)
    arrHeader(1) = "라우팅명"					' Header명(1)
    arrHeader(2) = "BOM Type"					' Header명(1)
    arrHeader(3) = "주라우팅"					' Header명(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
    frm1.txtRoutNo.focus
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value		= arrRet(0)		
		frm1.txtRoutNoDesc.Value		= arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
End Function


'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : OprNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	
	
	If frm1.txtRoutNo.value= "" Then
		Call DisplayMsgBox("971012","X", "라우팅","X")
		frm1.txtRoutNo.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
				  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "공정"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "공정"		
	arrHeader(1) = "작업장"	
	arrHeader(2) = "작업장명"
	arrHeader(3) = "공정작업명"
	arrHeader(4) = "사내구분"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "검사여부"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtOprNo.focus
	If arrRet(0) <> "" Then
		frm1.txtOprNo.Value	= arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(3)
	End If	
	
	Call SetFocusToDocument("M")
End Function

'=============================================  2.5.1 LoadInspCondition()  ======================================
'=	Event Name : LoadInspCondition
'=	Event Desc :
'========================================================================================================
Function LoadInspCondition()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then Exit Function
	End If
	
	With frm1
		
		.vspdData.Col = C_InspMthdCd
		.vspdData.Row = .vspdData.ActiveRow
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
			WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		End if
		
		If .vspdData.MaxRows  > 0 then 
			If Mid(Trim(.vspdData.Text), 2, 1) =  "2" Or Mid(Trim(.vspdData.Text), 2, 1) =  "3" Or Mid(Trim(.vspdData.Text), 2, 1) =  "9" Then				'검사방식이 선별형, 조정형이면 Exit
				Call DisplayMsgBox("220710", "X", "X", "X")  
				Exit Function
			End If
		
			If .vspdData.ActiveRow > 0 then
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col = C_InspItemCd
				WriteCookie "txtInspItemCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspItemNm
				WriteCookie "txtInspItemNm", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdCd
				WriteCookie "txtInspMthdCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdNm
				WriteCookie "txtInspMthdNm", Trim(.vspdData.Text)
			End If
		End If
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadInspStand1()  ======================================
'=	Event Name : LoadInspStand1
'=	Event Desc :
'========================================================================================================
Function LoadInspStand1()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		.vspdData.Col = C_InspMthdCd
		.vspdData.Row = .vspdData.ActiveRow
		
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		If .vspdData.MaxRows  > 0 then 
		
			If Mid(Trim(.vspdData.Text), 2, 1) <>  "3" Then				'검사방식이 조정형이 아니면 Exit
				Call DisplayMsgBox("220905", "X", "X", "X")  
				Exit Function
			End If
		
			If .cboInspClassCd.Value = "P" Then						'검사분류가 공정검사이면 Exit
				Call DisplayMsgBox("220740", "X", "X", "X")  
				Exit Function	
			End If
		
			if .vspdData.ActiveRow > 0 then
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col = C_InspItemCd
				WriteCookie "txtInspItemCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspItemNm
				WriteCookie "txtInspItemNm", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdCd
				WriteCookie "txtInspMthdCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdNm
				WriteCookie "txtInspMthdNm", Trim(.vspdData.Text)
			End If
		End if
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.5.3 LoadInspStand2()  ======================================
'=	Event Name : LoadInspStand2
'=	Event Desc :
'========================================================================================================
Function LoadInspStand2()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		
				
		.vspdData.Col = C_InspMthdCd
		.vspdData.Row = .vspdData.ActiveRow
		
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
		WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
		WriteCookie "txtOprNo", Trim(.txtOprNo.value)
		WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		
		If .vspdData.MaxRows  > 0 then
			If Mid(Trim(.vspdData.Text), 2, 1) <>  "3" Then				'검사방식이 조정형이 아니면 Exit
				Call DisplayMsgBox("220905", "X", "X", "X")  
				Exit Function
			End If
		
			If .cboInspClassCd.Value <> "P" Then						'검사분류가 공정검사가 아니면 Exit
				Call DisplayMsgBox("220708", "X", "X", "X") 
				Exit Function	
			End If
				
			if .vspdData.ActiveRow > 0 then
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col = C_InspItemCd
				WriteCookie "txtInspItemCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspItemNm
				WriteCookie "txtInspItemNm", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdCd
				WriteCookie "txtInspMthdCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdNm
				WriteCookie "txtInspMthdNm", Trim(.vspdData.Text)
			End If
		End if
	End With
	PgmJump(BIZ_PGM_JUMP3_ID)
End Function

'=============================================  2.5.2 LoadInspStand3()  ======================================
'=	Event Name : LoadInspStand3
'=	Event Desc :
'========================================================================================================
Function LoadInspStand3()
	Dim intRetCD
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
				
		.vspdData.Col = C_InspMthdCd
		.vspdData.Row = .vspdData.ActiveRow
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
			WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		End if
		
		If .vspdData.MaxRows  > 0 then
			If Mid(Trim(.vspdData.Text), 2, 1) <>  "2" Then				'검사방식이 선별형이 아니면 Exit
				Call DisplayMsgBox("220709", "X", "X", "X")  
				Exit Function
			End If
		
			if .vspdData.ActiveRow > 0 then
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col = C_InspItemCd
				WriteCookie "txtInspItemCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspItemNm
				WriteCookie "txtInspItemNm", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdCd
				WriteCookie "txtInspMthdCd", Trim(.vspdData.Text)
				.vspdData.Col = C_InspMthdNm
				WriteCookie "txtInspMthdNm", Trim(.vspdData.Text)
			End If
		End if
	End With
	
	PgmJump(BIZ_PGM_JUMP4_ID)
End Function

'=============================================  2.5.1 LoadInspStandCopy()  ======================================
'=	Event Name : LoadInspStandCopy
'=	Event Desc :
'========================================================================================================
Function LoadInspStandCopy()
	Dim intRetCD
	
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		'If .vspdData.MaxRows  = 0 then Exit Function
		
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
			WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		End if
		
	End With
	
	PgmJump(BIZ_PGM_JUMP5_ID)
End Function

'=============================================  2.5.1 CheckMinMax()  ======================================
'=	Event Name : CheckMinMax
'=	Event Desc : 입력값의 대소값 체크 
'========================================================================================================
Function CheckMinMax()
	Dim i
	Dim LSL, USL
	DIm LCL, UCL
	
	CheckMinMax = False
	
	With frm1
		For i = 0 To .vspdData.MaxRows
			If GetSpreadText(.vspdData,0,i,"X","X") = ggoSpread.InsertFlag or _
			   GetSpreadText(.vspdData,0,i,"X","X") = ggoSpread.UpdateFlag Then
				
				LSL = GetSpreadText(.vspdData,C_LSL,i,"X","X")
				USL = GetSpreadText(.vspdData,C_USL,i,"X","X")
				If LSL <> "" and USL <> "" Then
					If UNICDbl(LSL) >= UNICDbl(USL) Then
						Call DisplayMsgBox("800443", "X", "상한규격", "하한규격")  
						Call SetActiveCell(.vspdData, C_USL, i,"M","X","X")
						Exit Function
					End If
				End If

				LCL = GetSpreadText(.vspdData,C_LCL,i,"X","X")
				UCL = GetSpreadText(.vspdData,C_UCL,i,"X","X")
				If LCL <> "" and UCL <> "" Then
					If UNICDbl(LCL) >= UNICDbl(UCL) Then
						Call DisplayMsgBox("800443", "X", "관리상한", "관리하한")  
						Call SetActiveCell(.vspdData, C_UCL, i,"M","X","X")
						Exit Function
					End If
				End If
			End If
		Next
	End With
	CheckMinMax = True

End Function

'=============================================  2.5.1 CheckInspMthd_InspUnitIndctn()  ======================================
'=	Event Name : CheckInspMthd_InspUnitIndctn
'=	Event Desc :
'========================================================================================================
Function CheckInspMthd_InspUnitIndctn(Byval Row)

	Dim lngStartRow
	Dim lngEndRow
	Dim i
	
	CheckInspMthd_InspUnitIndctn = False
	With frm1.vspdData
		If Row = -1 then
			lngStartRow = 1
			lngEndRow = .MaxRows
		Else
			lngStartRow = Row
			lngEndRow = Row
		End If
		
		For i = lngStartRow To lngEndRow
			.Row = i
			.Col = 0 
			If .Text = lgInsertFlag Or  .Text = lgUpdateFlag Then 
				.Col = C_InspCharCd
				If Trim(.Text) = "" Then
					'없는 경우(직접 검사항목코드를 입력한 경우)
					'**** Msg : 검사항목은 반드시 팝업을 통해 선택하십시오.
					Call DisplayMsgBox("220716", "X", "X", "X")  
					.Col = C_InspItemCd
					.Action = 0
					Exit Function
				ElseIf Trim(.Text) = "C" Then
					'품질특성이 정성인 경우 
					.Col = C_InspMthdCd
					If Trim(.Text) = "" Then
						'검사방식이 선택되지 않은 경우 
						'****** Msg 검사방식을 선택하십시오.
						Call DisplayMsgBox("220906", "X", "X", "X")  
						.Action = 0
						Exit Function
					ElseIf Left(Trim(.Text), 1) = "1" Then
						'계수형인 경우 
						.Col = C_InspUnitIndctnCd
						If Trim(.Text) = "3" Then
							'****** Msg 품질특성이 정성이면서 검사방식이 계수형인 경우, 검사단위 품질표시로 특성치를 사용할 수 없습니다.
							Call DisplayMsgBox("220717", "X", "X", "X")  
							.Col = C_InspUnitIndctnNm
							.Action = 0
							Exit Function
						End If
					ElseIf  Left(Trim(.Text), 1) = "2" Then
						'계량형인 경우 
						'****** Msg 품질특성이 정성이면 검사방식을 계량형으로 선정할 수 없습니다.
						Call DisplayMsgBox("220718", "X", "X", "X")  
						.Action = 0
						Exit Function
					Else
						'그 외 
						.Col = C_InspUnitIndctnCd
						If Trim(.Text) = "3" Then
							'****** Msg 품질특성이 정성이면서 검사방식이 계수형이나 계량형이 아닌 경우, 검사단위 품질표시로 특성치를 사용할 수 없습니다.
							Call DisplayMsgBox("220719", "X", "X", "X")  
							.Col = C_InspUnitIndctnNm
							.Action = 0
							Exit Function
						End If
					End If
				ElseIf Trim(.Text) = "V" Then
					'품질특성이 정량인 경우 
					.Col = C_InspMthdCd
					If Left(Trim(.Text), 1) = "1" Then
						'계수형인 경우 
						.Col = C_InspUnitIndctnCd
						If Trim(.Text) = "2" Then
							'****** Msg 품질특성이 정량이면서 검사방식이 계수형인 경우 검사단위 품질표시로 결점수를 사용할 수 없습니다.
							Call DisplayMsgBox("220720", "X", "X", "X")  
							.Col = C_InspUnitIndctnNm
							.Action = 0
							Exit Function
						End If
					ElseIf  Left(Trim(.Text), 1) = "2" Then
						'계량형인 경우 
						.Col = C_InspUnitIndctnCd
						If Trim(.Text) <> "3" Then
							'****** Msg 품질특성이 정량이면서 검사방식이 계량형인 경우, 검사단위 품질표시로 특성치만 사용할 수 있습니다.
							Call DisplayMsgBox("220721", "X", "X", "X")  
							.Col = C_InspUnitIndctnNm
							.Action = 0
							Exit Function
						End If
					End If
				End If
			End If
		Next
	End With
	
	CheckInspMthd_InspUnitIndctn = True
End Function

'=============================================  2.5.2 CheckInspOrder()  ======================================
'=	Event Name : CheckInspOrder
'=	Event Desc :
'========================================================================================================
Function CheckInspOrder()
	Dim i
	Dim j
	Dim intInspOrder
	
	CheckInspOrder = False
	With frm1.vspdData
		
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				.Col = C_InspOrder
				intInspOrder = .Text
				For j = 1 To .MaxRows
					If i <> j Then
						.Row = j
						.Col = 0 
						If .Text <> ggoSpread.DeleteFlag Then
							.Col = C_InspOrder 
							If intInspOrder = .Text Then
								'**** Msg : 검사순서가 중복되었습니다.
								Call DisplayMsgBox("220723", "X", "X", "X")  
								.Action = 0
								Exit Function
							End If
						End If
					End If
				Next
			End If
		Next
	End With
	
	CheckInspOrder = True
End Function

'=============================================  2.5.2 CheckInspSpec()  ======================================
'=	Event Name : CheckInspSpec
'=	Event Desc :
'========================================================================================================
Function CheckInspSpec()
	Dim i
	Dim j
	Dim dblInspSpec
	
	CheckInspSpec = False
	With frm1.vspdData
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0 
			If .Text = lgInsertFlag Or  .Text = lgUpdateFlag Then 
				.Col = C_InspUnitIndctnCd
				Select Case .Text
					Case "2"
						.Col = C_InspSpec
						If Trim(.Text) = "" Then
							'**** Msg : 검사단위 품질표시가 결점수일 경우 검사규격에 허용결점수를 입력해야 합니다.
							Call DisplayMsgBox("220725", "X", "X", "X") 
							.Action = 0 	
							Exit Function
						Else
							On Error Resume Next
							dblInspSpec = UNICDbl(.Text)
							If Err.Number <> 0 then
								'**** Msg : 검사단위 품질표시가 결점수나 특성치일 경우 검사규격에 수치를 입력해야 합니다.
								Call DisplayMsgBox("220724", "X", "X", "X")  	
								.Action = 0
								Exit Function
							End If
							Err.Clear
							On Error GoTo 0
						End If
					Case "3"
						.Col = C_InspSpec
						If Trim(.Text) = "" Then
							'**** Msg : 검사단위 품질표시가 특성치일  경우 검사규격을 입력해야 합니다.
							Call DisplayMsgBox("220726", "X", "X", "X")  	
							.Action = 0
							Exit Function
						Else
							On Error Resume Next
							dblInspSpec = UNICDbl(.Text)
							If Err.Number <> 0 then
								'**** Msg : 검사단위 품질표시가 결점수나 특성치일 경우 검사규격에 수치를 입력해야 합니다.
								Call DisplayMsgBox("220724", "X", "X", "X")  	
								.Action = 0
								Exit Function
							End If
							Err.Clear
							On Error Goto 0
						End If
						
						.Col = C_LSL
						If Trim(.Text) = "" Then
							.Col = C_USL
							If Trim(.Text) = "" Then
								'**** Msg : 검사단위 품질표시가 특성치일 경우 적어도 한쪽규격을 입력해야 합니다.
								Call DisplayMsgBox("220727", "X", "X", "X")  	
								.Col = C_LSL
								.Action = 0
								Exit Function
							End If
						End IF
				End Select
			End If
		Next
	End With
	
	CheckInspSpec = True
End Function


'============================================= EnableField()  ======================================
'=	Event Name : EnableField
'=	Event Desc :
'========================================================================================================
Sub EnableField(Byval strInspClass)
	If	strInspClass = "P" Then
		Process.style.display	= ""
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "N")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "N")
	Else	
		Process.style.display	= "none"
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
	End if
End Sub

'============================================= cboInspClassCd_onchange()  ======================================
'=	Event Name : cboInspClassCd_onchange()
'=	Event Desc :
'========================================================================================================
Sub cboInspClassCd_onchange()
	Call EnableField(frm1.cboInspClassCd.value)
End Sub


'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
        
    gMouseClickStatus = "SPC"
 	
	Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
      
    Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
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

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolBar("11101101001011")							'⊙: 버튼 툴바 제어 
	Call EnableField(frm1.cboInspClassCd.value)	
    	If Trim(frm1.txtPlantCd.Value) = "" Then
    		frm1.txtPlantCd.focus 
    	Else
    		frm1.txtItemCd.focus
		End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	
	With frm1.vspdData
		If Col = C_InspOrder Then
			.Col = Col
			.Row = Row
			If .Text = 0 Then
				Call DisplayMsgBox("141704", "X", "검사순서", "X")  
				.Text = ""
			End If
		End If
	End With
	
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
			Case  C_WeightNm
				.Col = Col
				intIndex = .Value
				.Col = C_WeightCd
				.Value = intIndex
			Case  C_MthdOfCLCalNm
				.Col = Col
				intIndex = .Value
				.Col = C_MthdOfCLCalCd
				.Value = intIndex
			Case  C_InspUnitIndctnNm
				.Col = Col
				intIndex = .Value
				.Col = C_InspUnitIndctnCd
				.Value = intIndex
				Call CheckInspMthd_InspUnitIndctn(Row)
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
		
		If Row > 0 And Col = C_InspItemPopUp Then
			.Col = C_InspItemCd
			.Row = Row
			        
			Call OpenInspItem(.Text)
			
		ElseIf Row > 0 And Col = C_InspMthdPopUp Then
			.Col = C_InspMthdCd
			.Row = Row
			
			Call OpenInspMthd(.Text)    
		
		ElseIf Row > 0 And Col = C_MeasmtEquipmtPopUp Then
			.Col = C_MeasmtEquipmtCd
			.Row = Row
			
			Call OpenMeasmtEquipmt(.Text)    
		
		ElseIf Row > 0 And Col = C_MeasmtUnitPopUp Then
			.Col = C_MeasmtUnitCd
			.Row = Row
					
			Call OpenMeasmtUnit(.Text)            
		
		End If
	
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData
	
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
End Sub 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
	
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear                                                            		   '☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
        If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then	Exit Function
    	End If
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then	Exit Function
	
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")  
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call InitVariables
	
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function
	
	FncQuery = True						'⊙: Processing is OK
    
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
		If IntRetCD = vbNo Then	Exit Function
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
	Call EnableField(frm1.cboInspClassCd.value)
	
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
	Dim IntRetCD 
	
	FncDelete = False                                                       '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	'On Error Resume Next                                                    '☜: Protect system from crashing
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900005", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False                                                  		       '⊙: Processing is NG

	Err.Clear                                                            	 		  '☜: Protect system from crashing
	
	On Error Resume Next                                           	       '☜: Protect system from crashing
	   
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
	If Not chkField(Document, "1") Then Exit Function
    	
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSDefaultCheck = False Then Exit Function

    '상하한 입력값 대소 체크 
    If CheckMinMax() = False Then Exit Function

    '검사순서 중복 체크 
    If CheckInspOrder() = False Then Exit Function
    	
    '검사방법 및 검사단위 품질 표시 적합성 체크 
    If CheckInspMthd_InspUnitIndctn(-1) = False Then Exit Function
    	
    '검사단위 품질 표시에 맞는 Spec 작성 여부 체크 
    If CheckInspSpec = False Then Exit Function
    	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function                              '☜: Save db data
    
	FncSave = True                                      	                    '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	
	With frm1.vspdData
		
		If .MaxRows < 1 then Exit function
		
		.ReDraw = False
		
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
	    
	    .Row = .ActiveRow
	    .Col = C_InspItemCd
	    .Text = ""
	    .Col = C_InspItemNm
	    .Text = ""


		Dim i
		Dim iMaxCount
		Dim iUpdateFlag
		iMaxCount = 0
		
		If .MaxRows <> 1 Then
			.Col = C_InspOrder
			for i = 1 to .MaxRows
				.Row = i
				If CInt(.Text) > CInt(iMaxCount) Then
					iMaxCount = CInt(.Text)
				End IF
			Next
		
			.Row = .ActiveRow
			.Col = C_InspOrder
			.Text = CInt(iMaxCount) + 1
		End If	    	
	    	
	    .ReDraw = True                                   					            '☜: Protect system from crashing
	    
		Call SetActiveCell(frm1.vspdData,C_InspItemCd,.ActiveRow,"M","X","X")
		Set gActiveElement = document.ActiveElement	    
	End With
	FncCopy = true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	FncCancel = false
	If frm1.vspdData.MaxRows < 1 then Exit function
    ggoSpread.Source = frm1.vspdData	
	ggoSpread.EditUndo   
	FncCancel = true                                               '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
	
	'/* 2003-06 정기패치: 행이 하나도 없을 경우에 여러 행을 추가하면 검사순서가 2번부터 채번되는 것 수정 - START */
	Dim NoRowFlag 
	
	On Error Resume Next
	 
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
	End If
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		'.EditMode = True
		.ReDraw = False
		
		If .MaxRows = 0 Then
			NoRowFlag = True
		End if
		
		ggoSpread.InsertRow .ActiveRow, imRow
		
		Dim i
		Dim iMaxCount
		Dim iUpdateFlag
		iMaxCount = 0
		
		If imRow = 1 Then
			.Col = C_InspOrder
			
			'/* 검사순서에 순번이 잘못 채번되는 오류 수정(2005-09-28) - START */
			If NoRowFlag = False Then
				for i = 1 to .MaxRows
					.Row = i
					If CInt(.Value) > iMaxCount Then
						iMaxCount = CInt(.Value)
					End IF
				Next	
			End If
			
			.Row = .ActiveRow
			.Value = CInt(iMaxCount) + 1
			'/* 검사순서에 순번이 잘못 채번되는 오류 수정(2005-09-28) - START */			
		Else
			.Col = C_InspOrder
			'/* 검사순서에 순번이 잘못 채번되는 오류 수정(2005-09-28) - START */
			If NoRowFlag = False Then
				for i = 1 to .MaxRows
					.Row = i
					If CInt(.Value) > iMaxCount Then
						iMaxCount = CInt(.Value)
					End IF
				Next
			End If				
			
			for i = 0 to imRow - 1
				.Row = .ActiveRow + i
				iMaxCount = iMaxCount + 1
				.Value = iMaxCount
			Next
			'/* 검사순서에 순번이 잘못 채번되는 오류 수정(2005-09-28) - END */
			
		End If
		'/* 2003-06 정기패치: 행이 하나도 없을 경우에 여러 행을 추가하면 검사순서가 2번부터 채번되는 것 수정 - END */

		SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
		
		.ReDraw = True
		
    
		Call SetActiveCell(frm1.vspdData,C_InspItemCd,.ActiveRow,"M","X","X")
		Set gActiveElement = document.ActiveElement
    
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
    
	If frm1.vspdData.MaxRows < 1 then Exit function
	
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData 
	
	lDelRows = ggoSpread.DeleteRow
	
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    FncNext = false                                                  '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

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
		If IntRetCD = vbNo Then Exit Function
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Err.Clear                                                               			'☜: Protect system from crashing
	
	Call LayerShowHide(1)
	
	DbQuery = False
		
	With frm1	
			
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode="			& Parent.UID_M0001 _
									& "&txtPlantCd="		& .hPlantCd.value _
									& "&txtItemCd="			& .hItemCd.value _
									& "&cboInspClassCd="	& .hInspClassCd.value _
									& "&txtRoutNo="			& .hRoutNo.value _
									& "&txtOprNo="			& .hOprNo.value _
									& "&lgStrPrevKey="		& lgStrPrevKey _
									& "&txtMaxRows="		& .vspdData.MaxRows			
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode="			& Parent.UID_M0001 _
									& "&txtPlantCd="		& Trim(.txtPlantCd.Value) _
									& "&txtItemCd="			& Trim(.txtItemCd.value) _
									& "&cboInspClassCd="	& Trim(.cboInspClassCd.Value) _
									& "&txtRoutNo="			& Trim(.txtRoutNo.value) _
									& "&txtOprNo="			& Trim(.txtOprNo.value) _
									& "&lgStrPrevKey="		& lgStrPrevKey _
									& "&txtMaxRows="		& .vspdData.MaxRows			
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)					'☜: 비지니스 ASP 를 가동 
				
		DbQuery = True                                                          			'⊙: Processing is NG
	End With

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()					'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			'⊙: Indicates that current mode is Update mode
	Call SetToolBar("11101111001111")							'⊙: 버튼 툴바 제어 
	
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	Call EnableField(frm1.cboInspClassCd.value)
	
	Dim posActiveRow
	If frm1.vspdData.MaxRows <= 100 Then
		posActiveRow = 1
	ElseIf frm1.vspdData.MaxRows > 100 Then
		posActiveRow = frm1.vspdData.MaxRows - 99
	End If
	
	Call SetActiveCell(frm1.vspdData,C_InspOrder,posActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpInsCnt
	Dim lGrpDelCnt 
	Dim strDel
	Dim strVal

	Dim iColSep
	Dim iRowSep
	Dim iInsertFlag
	Dim iUpdateFlag
	Dim iDeleteFlag
	Dim arrVal
	Dim arrDel

	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing
	
	iColSep     = Parent.gColSep
	iRowSep     = Parent.gRowSep
	iInsertFlag = ggoSpread.InsertFlag
	iUpdateFlag = ggoSpread.UpdateFlag
	iDeleteFlag = ggoSpread.DeleteFlag                                                   '☜: Protect system from crashing

	With frm1
	
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpInsCnt = 0
		lGrpDelCnt = 0 
		
		ReDim arrVal(0)
		ReDim arrDel(0)

	   	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    		.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
				Case iInsertFlag					'☜: 신규 
					
					strVal = "C" & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspOrder,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspMthdCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspUnitIndctnCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_WeightCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspSpec,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_LSL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_USL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MthdOfCLCalCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_CalculatedQty,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_LCL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_UCL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtEquipmtCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtUnitCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspProcessDesc,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_Remark,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep
					
					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1
					
				Case iUpdateFlag					'☜: 수정 
					
					strVal = "U" & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspOrder,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspMthdCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspUnitIndctnCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_WeightCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspSpec,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_LSL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_USL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MthdOfCLCalCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_CalculatedQty,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_LCL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_UCL,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtEquipmtCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_MeasmtUnitCd,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_InspProcessDesc,lRow,"X","X") & iColSep _
								 & GetSpreadText(.vspdData,C_Remark,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrVal(lGrpInsCnt)
					arrVal(lGrpInsCnt) = strVal
					lGrpInsCnt = lGrpInsCnt + 1

				Case iDeleteFlag					'☜: 삭제 
					
					strDel = "D" & iColSep _
								 & GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & iColSep _
								 & CStr(lRow) & iRowSep

					ReDim Preserve arrDel(lGrpDelCnt)
					arrDel(lGrpDelCnt) = strDel
					lGrpDelCnt = lGrpDelCnt + 1

			End Select
		Next
		
		strVal = Join(arrVal,"")
		strDel = Join(arrDel,"")
		
		.txtMaxRows.value = lGrpInsCnt + lGrpDelCnt
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'☜: 비지니스 ASP 를 가동 
	End With
	
	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()				            '☆: 저장 성공후 실행 로직 
	DbSaveOk = false

	With frm1
		.vspdData.MaxRows = 0

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			.txtPlantCd.Value	= .hPlantCd.value
			.txtItemCd.Value	= .hItemCd.value
			.txtRoutNo.value	= .hRoutNo.value
			.txtOprNo.value		= .hOprNo.value
			.cboInspClassCd.Value = .hInspClassCd.value
		End If
	End With

   	Call InitVariables
    Call MainQuery()
	DbSaveOk = true
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	DbDelete = false
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별 검사기준 등록</font></td>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWPAP>공장</TD>
									<TD CLASS="TD6" NOWPAP>
										<input TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU" ><IMG align=top height=20 name=btnPlantCd1 onclick=vbscript:OpenPlant() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtPlantNm" SIZE="20" tag="14" >
									</TD>
									<TD CLASS="TD5" NOWPAP>검사분류</TD>
									<TD CLASS="TD6" NOWPAP>
										<SELECT Name="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWPAP>품목</TD>
									<TD CLASS="TD6" NOWPAP>
										<input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="품목" tag="12XXXU" ><IMG align=top height=20 name=btnItemCd1 onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="20" tag="14" >
									</TD>
									<TD CLASS="TD5" NOWPAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWPAP>&nbsp;</TD>										
								</TR>
								<TR ID="Process">
					      			<TD CLASS="TD5" NOWPAP>라우팅</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="12XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWPAP>공정</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
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
						<TABLE WIDTH="100%" HEIGHT="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/q1211ma1_I383095788_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
   		     			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspStand1">조정형(공정 외) 검사조건</A>&nbsp;|&nbsp;<A href="vbscript:LoadInspStand2">조정형(공정) 검사조건</A>&nbsp;|&nbsp;<A href="vbscript:LoadInspStand3">선별형 검사조건</A>&nbsp;|&nbsp;<A href="vbscript:LoadinspCondition">기타검사조건</A>&nbsp;|&nbsp;<A href="vbscript:LoadInspStandCopy">검사기준 복사</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
	    			</TR>
	    		</TABLE>
	    	</TD>
         </TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1 ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24" tabindex=-1 >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
