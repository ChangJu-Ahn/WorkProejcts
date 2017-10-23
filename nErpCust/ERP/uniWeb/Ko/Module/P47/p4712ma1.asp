
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Adjust Resource Compution
'*  3. Program ID           : p4712ma1
'*  4. Program Name         : Adjust Resource Compution
'*  5. Program Desc         : Adjust Resource Compution
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/30
'*  8. Modified date(Last)  : 2002/07/18
'*  9. Modifier (First)     : Jeon, JaeHyun
'* 10. Modifier (Last)      : Kang Seong Moon
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_LOOKUP_ID	= "p4712mb0.asp"								' Head의 제조오더의 공정에 대한 정보 

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p4712mb1.asp"								'☆: BOR(Bill Of Resource) 정보 조회 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p4712mb2.asp"								'☆: Main Query(자원실적정보) 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "p4712mb3.asp"                                '☆: 자원소비실적 처리 
Const BIZ_PGM_LOOKUPRC_ID	= "p4712mb4.asp"                            '☆: 자원정보 LOOKUP 처리 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation 
Dim C_ResourceCd			'= 1
Dim C_ResourcePopup			'= 2
Dim C_ResourceNm			'= 3
Dim C_ConsumedDt			'= 4
Dim C_ConsumedQty			'= 5
Dim C_ResourceTypeNm		'= 6
Dim C_ResourceGroupCd		'= 7
Dim C_ResourceGroupNm		'= 8


' Grid 2(vspdData2) - Operation
Dim C_ResourceCd2			'= 1
Dim C_ResourceNm2			'= 2
Dim C_ResourceTypeNm2		'= 3
Dim C_ResourceGroupCd2		'= 4
Dim C_ResourceGroupNm2		'= 5
Dim	C_Rank2					'= 6
Dim	C_BOR_Efficiency2		'= 7
Dim C_ValidFromDt			'= 8
Dim C_ValidToDt				'= 9

Dim strDate
Dim BaseDate
Dim strYear
Dim strMonth
Dim strDay

BaseDate = "<%=GetSvrDate%>"

Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)			'☆: 초기화면에 뿌려지는 날짜 

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size를 조사할 변수 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey

Dim lgLngCurRows
Dim lgCurrRow
Dim lgFlgQueryCnt
Dim lgSortKey
Dim lgSortKey2
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgRow         
'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgRow = 0
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=================================================================================================== 
Sub SetCookieVal()
	
	frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
	frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
	frm1.txtItemCd.Value	= ReadCookie("txtItemCd")

	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement

		
	frm1.txtProdOrderNo.Value	= ReadCookie("txtProdOrderNo")
	frm1.txtOprCd.Value			= ReadCookie("txtOprNo")
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtProdOrderNo", ""
	WriteCookie "txtOprNo", ""
		
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call InitSpreadPosVariables(pvSpdNo)
	
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
		
	if pvSpdNo = "*" or pvSpdNo = "A" then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_ResourceGroupNm +1													'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_ResourceCd,		"자원코드",	10,,,10,2
		ggoSpread.SSSetButton 	C_ResourcePopup
		ggoSpread.SSSetEdit		C_ResourceNm,		"자원명",	20
		ggoSpread.SSSetDate		C_ConsumedDt,		"자원소비일",	13,	2,	parent.gDateFormat
		ggoSpread.SSSetTime		C_ConsumedQty,		"자원소비시간",	13,2 ,1 ,1
		ggoSpread.SSSetEdit		C_ResourceTypeNm,	"자원구분",	10
		ggoSpread.SSSetEdit		C_ResourceGroupCd,	"자원그룹",	10
		ggoSpread.SSSetEdit		C_ResourceGroupNm,	"자원그룹명",	20
		
		Call ggoSpread.MakePairsColumn(C_ResourceCd,C_ResourcePopup)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(4)
		.ReDraw = true
		End With
	end if
	
	if pvSpdNo = "*" or pvSpdNo = "B" then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_ValidToDt +1										'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit		C_ResourceCd2,		"자원코드", 10
		ggoSpread.SSSetEdit		C_ResourceNm2,		"자원명",	20
		ggoSpread.SSSetEdit		C_ResourceTypeNm2,	"자원구분", 10
		ggoSpread.SSSetEdit		C_ResourceGroupCd2, "자원그룹", 10
		ggoSpread.SSSetEdit		C_ResourceGroupNm2, "자원그룹명", 20
		ggoSpread.SSSetFloat	C_Rank2,			"순서",		10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BOR_Efficiency2,	"효율",		10, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate		C_ValidFromDt,		"시작일",	11,	2,	parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt,		"종료일",	11,	2,	parent.gDateFormat
		'Call ggoSpread.MakePairsColumn(,)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(2)	
		.ReDraw = true
		End With
    end if
	
	Call SetSpreadLock 
    
End Sub


'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock()

    With frm1
    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = frm1.vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.SpreadLock C_ResourceCd,-1,C_ResourceCd
	ggoSpread.SpreadLock C_ResourcePopup,-1,C_ResourcePopup
	ggoSpread.SpreadLock C_ResourceNm,-1,C_ResourceNm
    ggoSpread.SpreadLock C_ConsumedDt,	-1
	ggoSpread.SpreadLock C_ResourceTypeNm, -1
	ggoSpread.SpreadLock C_ResourceGroupCd, -1
	ggoSpread.SpreadLock C_ResourceGroupNm, -1
	ggoSpread.SpreadLock frm1.vspdData1.MaxCols, -1, frm1.vspdData1.MaxCols
	
	ggoSpread.SpreadUnLock  C_ConsumedQty, -1 ,C_ConsumedQty
	ggoSpread.SSSetRequired  C_ConsumedQty, -1 , C_ConsumedQty
	.vspdData1.ReDraw = True
    
    '--------------------------------
    'Grid 2
    '--------------------------------
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()

    End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1.vspdData1
    
    .Redraw = False
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSSetRequired  C_ResourceCd,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceNm,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_ConsumedDt,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_ConsumedQty,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceTypeNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceGroupCd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceGroupNm,		pvStartRow, pvEndRow

    .Redraw = True
    
    End With

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
'Sub InitComboBox()

'End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	if pvSpdNo = "*" or pvSpdNo = "A" then
		'Grid 1(vspdData1) - Operation 
		C_ResourceCd			= 1
		C_ResourcePopup			= 2
		C_ResourceNm			= 3
		C_ConsumedDt			= 4
		C_ConsumedQty			= 5
		C_ResourceTypeNm		= 6
		C_ResourceGroupCd		= 7
		C_ResourceGroupNm		= 8
	end if

	if pvSpdNo = "*" or pvSpdNo = "B" then
		'Grid 2(vspdData2) - Operation
		C_ResourceCd2			= 1
		C_ResourceNm2			= 2
		C_ResourceTypeNm2		= 3
		C_ResourceGroupCd2		= 4
		C_ResourceGroupNm2		= 5
		C_Rank2					= 6
		C_BOR_Efficiency2		= 7
		C_ValidFromDt			= 8
		C_ValidToDt				= 9
	end if

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
  	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
  	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
  		
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ResourceCd			= iCurColumnPos(1)
		C_ResourcePopup			= iCurColumnPos(2)
		C_ResourceNm			= iCurColumnPos(3)
		C_ConsumedDt			= iCurColumnPos(4)
		C_ConsumedQty			= iCurColumnPos(5)
		C_ResourceTypeNm		= iCurColumnPos(6)
		C_ResourceGroupCd		= iCurColumnPos(7)
		C_ResourceGroupNm		= iCurColumnPos(8)
	Case "B"
 		ggoSpread.Source = frm1.vspdData2
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ResourceCd2			= iCurColumnPos(1)
		C_ResourceNm2			= iCurColumnPos(2)
		C_ResourceTypeNm2		= iCurColumnPos(3)
		C_ResourceGroupCd2		= iCurColumnPos(4)
		C_ResourceGroupNm2		= iCurColumnPos(5)
		C_Rank2					= iCurColumnPos(6)
		C_BOR_Efficiency2		= iCurColumnPos(7)
		C_ValidFromDt			= iCurColumnPos(8)
		C_ValidToDt				= iCurColumnPos(9) 
  	End Select
  
End Sub
 

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
		
End Function

'------------------------------------------  OpenOprCd()  -------------------------------------------------
'	Name : OpenOprCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprCd()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtOprCd.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtProdOrderNo.value = "" Then
		Call DisplayMsgBox("971012","X", "제조오더번호","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdOrderNo.value
	arrParam(2) = "Y"
	
	iCalledAspName = AskPRAspName("p4112pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4112pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetOprCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource(ByVal strCode,ByVal strName,ByVal Row)

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	arrParam(5) = "자원"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원"		
    arrHeader(1) = "자원명"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
		
End Function

'------------------------------------------  OpenPartRef()  -------------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4311ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4311ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4111ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'---------------------------------------------------------------------------------------------------------
Function OpenProdRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 

   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
		arrParam(2) = .hOprNo.Value
	End With
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4411ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4411ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
    With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4511ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4511ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P45/p4511ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConsumRef()  -------------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'---------------------------------------------------------------------------------------------------------
Function OpenConsumRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
    With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4412ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4412ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P44/p4412ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetOprCd()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprCd(byval arrRet)
	frm1.txtOprCd.Value    = arrRet(0)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	With frm1.vspdData1
		ggoSpread.Source = frm1.vspdData1
			.Row = frm1.vspdData1.ActiveRow
			.Col = C_ResourceCd
			.Text = arrRet(0)
			
			'.Col = C_ResourceNm
			'.Text = arrRet(1)
			LookUpRc arrRet(0), .Row
	End with		
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 


Sub InitData(ByVal lngStartRow)

End Sub

'==============================================================================
' Function : ConvToSec()
' Description : 저장시에 각 시간 데이터들을 초로 환산 
'==============================================================================
Function ConvToSec(ByVal Str)

	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

'==========================================  2.5.2 LookupRc()  =======================================
'	Name : LookUpRc()
'	Description : Lookup WorkCenter using Keyboard
'===================================================================================================== 
Sub LookUpRc(ByVal strResourceCd, ByVal Row)
	Dim strVal		
	Call LayerShowHide(1)

	strVal = BIZ_PGM_LOOKUPRc_ID & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtResourceCd=" & Trim(strResourceCd)								'☜: 조회 조건 데이타 
	strVal = strVal & "&Row=" & Row											'☜: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Sub

Function LookUpRcOk(ByVal Rcd, ByVal RcdNm,ByVal RcdType, ByVal Rgcd, Byval RgcdNm, Byval Row)

	With frm1.vspdData1
				
		ggoSpread.Source = frm1.vspdData1
				
		.Row = Row
		.Col = C_ResourceNm
		.Text = RcdNm
		
		.Col = C_ResourceTypeNm
		.Text = RcdType
			
		.Col = C_ResourceGroupCd
		.Text = Rgcd
		
		.Col = C_ResourceGroupNm
		.Text = RgcdNm	
				
	End With

	IsOpenPop = False
		
End Function

Function LookUpRcNotOk(Byval Row)

	With frm1.vspdData1
				
		ggoSpread.Source = frm1.vspdData1
		
		.Row = Row
		.Col = C_ResourceCd
		.Text = ""
		.Col = C_ResourceNm
		.Text = ""
		.Col = C_ResourceTypeNm
		.Text = ""
		.Col = C_ResourceGroupCd
		.Text = ""
		.Col = C_ResourceGroupNm
		.Text = ""
	End With

	IsOpenPop = False
		
End Function


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)        
	    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    
	If ReadCookie("txtPlantCd") <> "" Then
		Call SetCookieVal
		frm1.txtOprCd.focus 
		Set gActiveElement = document.activeElement 
	Else	
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = UCase(parent.gPlant)
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtProdOrderNo.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================

'========================================================================================
' Function Name : vspdDat1a_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
  	Else
  		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	End If
  	
  	gMouseClickStatus = "SPC"   
     
  	Set gActiveSpdSheet = frm1.vspdData1
     
  	If frm1.vspdData1.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData1 
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
' Function Name : vspdData2_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
  	gMouseClickStatus = "SP2C"   
     
  	Set gActiveSpdSheet = frm1.vspdData2
    
    Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
     
  	If frm1.vspdData2.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData2 
  		If lgSortKey2 = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey2 = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
  			lgSortKey2 = 1
  		End If
 	Else
  		'------ Developer Coding part (Start)
 	 	'------ Developer Coding part (End)
 	
  	End If

End Sub
 
'=======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col, ByVal Row)

	Dim strItemCd
	Dim strHndItemCd, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim	DblRqrdQty
	Dim strResourceCd
	
	ggoSpread.Source = frm1.vspdData1
	
	If Row < 1 Then Exit Sub
	
	With frm1.vspdData1	
	.Row = Row

	Select Case Col
			
			Case C_ResourceCd
				ggoSpread.Source = frm1.vspdData1
				frm1.vspdData1.Col = Col	
				frm1.vspdData1.Row = Row
				strResourceCd = frm1.vspdData1.text
					If frm1.vspdData1.Text <> "" Then
						Call LookUpRc(strResourceCd, Row)
					End If

				IsOpenPop = True  '?
			

		    Case C_ConsumedQty
						
			    ggoSpread.Source = frm1.vspdData1
				ggoSpread.UpdateRow Row
			
		    Case C_ConsumedDt
		    
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.UpdateRow Row

		End Select

	End With

End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData1
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
     ggoSpread.Source = frm1.vspdData2
     Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
' 	If NewCol = C_XXX or Col = C_XXX Then
 '		Cancel = True
 '		Exit Sub
 '	End If
     ggoSpread.Source = frm1.vspdData1
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
 	 ggoSpread.Source = frm1.vspdData2
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("B")
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
    dim pvSpdNo
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)
    Select Case pvSpdNo
	case "A"
		ggoSpread.Source = frm1.vspdData1
	case "B"
		ggoSpread.Source = frm1.vspdData2
	End Select 
	Call ggoSpread.ReOrderingSpreadData()
	 
End Sub 

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

Dim strCode
Dim strName

    With frm1.vspdData1
    
		ggoSpread.Source = frm1.vspdData1
		If Row < 1 Then Exit Sub

		Select Case Col

		    Case C_ResourcePopup
				.Col = C_ResourceCd
				.Row = Row
				strCode = .Text
				.Col = C_ResourceNm
				.Row = Row
				strName = .Text
				Call OpenResource(strCode, strName, Row)
				
				Call SetActiveCell(frm1.vspdData1,C_ResourceCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement

		End Select

	End With

End Sub


'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If 
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1 ,NewTop) Then

		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub


Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If 
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then

		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData1_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub 
  
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'		========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SP2C" Then
        gMouseClickStatus = "SP2CR"
     End If
End Sub 

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False											'⊙: Processing is NG
    
    Err.Clear													'☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
        
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
              
    Call InitVariables
    frm1.vspdData1.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
	lgFlgQueryCnt = 0

    '-----------------------
    'Query function call area
    '-----------------------
    
    If DbQuery = False Then
		 Call RestoreToolBar()
		 Exit Function												'☜: Query db data
    End If
       
    FncQuery = True												'⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	
    Dim IntRetCD 
    Dim	LngRows
    
    FncSave = False												'⊙: Processing is NG
    
    Err.Clear                                                   '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    
    With frm1.vspdData1
     
    For LngRows = 1 To .MaxRows
      .Row = LngRows
      .Col = 0
      If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
		
		.Col = C_ConsumedQty					
			If .Text = "<%=ConvToTimeFormat(0)%>" Then
				Call DisplayMsgBox("189715", "x", "x", "x")
				Exit Function
			End If 
			
		.Col = C_ConsumedDt
			If CompareDateByFormat(.text,"<%=strDate%>","자원소비일","현재일","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
			  Exit Function               
			End If
	   End If
    Next
    
    End With
        
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function													'☜: Save db data
    
    FncSave = True												'⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
        
	If frm1.vspdData1.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData1.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData1.EditMode = True
	frm1.vspdData1.ReDraw = False
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.CopyRow
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
    frm1.vspdData1.Col = C_ResourceCd
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceNm
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceTypeNm
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceGroupCd
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceGroupNm
    frm1.vspdData1.Text = ""
	SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow
    frm1.vspdData1.ReDraw = True
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

  If frm1.vspdData1.MaxRows < 1 Then Exit Function	
  ggoSpread.EditUndo                                                  '☜: Protect system from crashing


End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 

    Dim IntRetCD
 	Dim imRow
 	Dim i
 	
 	On Error Resume Next
 	FncInsertRow = False

 	If IsNumeric(Trim(pvRowCnt)) Then
 		imRow = Cint(pvRowCnt)
 	Else
 		imRow = AskSpdSheetAddRowCount()
 		If imRow = "" Then
 			Exit Function
 		End If
 	End If
    
    With frm1
    	.vspdData1.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData1
			.vspdData1.ReDraw = False
	    ' 	ggoSpread.InsertRow
			ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
	     	SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1
			for i = 0 to imRow -1
				.vspdData1.Row = .vspdData1.ActiveRow + i
				.vspdData1.Col = C_ConsumedDt
				.vspdData1.text = strDate
				.vspdData1.Col = C_ConsumedQty
				.vspdData1.text = "<%=ConvToTimeFormat(0)%>"
				.vspdData1.ReDraw = True
			next
		'SetSpreadColor .vspdData1.ActiveRow		
	End With
    If Err.number = 0 Then FncInsertRow = True

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i

    With frm1

		'.vspdData1.Row = frm1.vspdData1.ActiveRow
		'.vspdData1.Col = C_OrderStatus2
    	'If .vspdData1.Text = "ST" or .vspdData1.Text = "CL" Then
		'	Call DisplayMsgBox("189520", "x", "x", "x")
	    '		Exit Function
		'End IF    
		If .vspdData1.MaxRows < 1 Then Exit Function

		'Call DeleteMarkingHSheet()

    End With

	ggoSpread.Source = frm1.vspdData1
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

	'CopyToHSheet frm1.vspdData1.ActiveRow

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.fncPrint()                                                   '☜: Protect system from crashing
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    
    lgFlgQueryCnt = lgFlgQueryCnt + 1
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_LOOKUP_ID	 & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtOprNo=" & Trim(.hOprCd.Value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows		
	Else
		strVal =  BIZ_PGM_LOOKUP_ID	 & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()

	lgOldRow = 1

	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement	
			If DbDtlQuery = False Then
				 Call RestoreToolBar()	
				 Exit Function 
			End If 
		End If
	End If

End Function


'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 실패일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryNotOk()

'frm1.txtPlantCd.focus


End Function


'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
    
	boolExist = False
    With frm1
    
	DbDtlQuery = False   
    
		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.Value)
			strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.Value)
			strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtlQuery = True

End Function



Function DbDtlQueryOk()												'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call DbDtl2Query
		End If
	End If

End Function


'========================================================================================
' Function Name : DbDtl2Query
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtl2Query() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
    
	boolExist = False
    With frm1
    
	DbDtl2Query = False   
    
		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtl2Query = True

End Function



Function DbDtl2queryOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	Call SetToolbar("11001111001111")										'⊙: 버튼 툴바 제어 
		
	frm1.vspdData1.ReDraw = False
   
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgAfterQryFlg = True
    Call InitData(LngMaxRow)
    frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1	
	frm1.vspdData1.ReDraw = True

End Function

Function DbDtl2querynotOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	Call SetToolbar("11001101001111")										'⊙: 버튼 툴바 제어 
		
	frm1.vspdData1.ReDraw = False
   
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgAfterQryFlg = True
    Call InitData(LngMaxRow)
    frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1	
	frm1.vspdData1.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData()
    
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================


Function DbSave() 

    Dim IntRows
    Dim strVal, strDel
	Dim ChkTimeVal
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size
		
    DbSave = False                                                          '⊙: Processing is NG
      
    Call LayerShowHide(1)

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0


	With frm1.vspdData1

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
	
			Select Case .Text
		    
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
		            
		            strVal = ""
		            
		            If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
					Else
						strVal = strVal & "U" & iColSep				'⊙: U=Update
					End If	
		            
					strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					strVal = strVal & UCase(Trim(frm1.txtProdOrderNo.value)) & iColSep
					strVal = strVal & UCase(Trim(frm1.txtOprCd.value)) & iColSep

					.Col = C_ResourceCd
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ConsumedDt		' ConsumedDt
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					
					.Col = C_ConsumedQty	' ConsumedQty
					ChkTimeVal = ConvToSec(Trim(.Text))
					
					If ChkTimeVal = -999999	Then
						Call DisplayMsgBox("970029", vbInformation, "자원소비시간", "", I_MKSCRIPT)
						Call SheetFocus(arrVal(1),8,I_MKSCRIPT)
						Response.End	
					Else
						strVal = strVal & ChkTimeVal & iRowSep
					End If

			    Case ggoSpread.DeleteFlag
			    
					strDel = "" 
					
					strDel = strDel & "D" & iColSep				'⊙: D=Delete
					strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					strDel = strDel & UCase(Trim(frm1.txtProdOrderNo.value)) & iColSep
					strDel = strDel & UCase(Trim(frm1.txtOprCd.value)) & iColSep

					.Col = C_ResourceCd
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_ConsumedDt		' ConsumedDt
					strDel = strDel & UNIConvDate(Trim(.Text)) & iRowSep

			End Select
			
			.Col = 0
			
			Select Case .Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
				Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
	
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                                           ' ⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    lgLngCurRows = 0                           'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
	
	Call RemovedivTextArea
	Call DbDtl2Query
	
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
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
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function

</script>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원소비등록(오더별)</font></td> 
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;<A href="vbscript:OpenPartRef()">부품내역</A> | <A href="vbscript:OpenOprRef()">공정내역</A> | <A href="vbscript:OpenProdRef()">실적내역</A> | <A href="vbscript:OpenRcptRef()">입고내역</A> | <A href="vbscript:OpenConsumRef()">부품소비내역</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
								    <TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
									<TD CLASS=TD5 NOWRAP>공정</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprCd" SIZE=8 MAXLENGTH=3 tag="12xxxU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprCd()"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>품목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="24" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>오더수량</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I549825119_txtOrderQty.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>작업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="24" ALT="작업장">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>오더단위</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtOrderUnit SIZE=10 MAXLENGTH=10 tag="24xxxU" ALT="오더단위"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Tracking No</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="24" ALT="Tracking No">&nbsp;</TD>
								<TD CLASS=TD5 NOWRAP>실적수량</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I176752752_txtProdQty.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>라우팅</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=18 MAXLENGTH=20 tag="24" ALT="라우팅">&nbsp;</TD>
								<TD CLASS=TD5 NOWRAP>양품수량</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I264298713_txtGoodQty.js'></script>
								</TD>
							</TR>	
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4712ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4712ma1_B_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtRoutNo" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
