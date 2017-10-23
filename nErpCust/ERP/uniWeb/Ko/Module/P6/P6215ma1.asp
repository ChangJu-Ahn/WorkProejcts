<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: 금형관리 
'*  2. Function Name		: 금형점검계획 
'*  3. Program ID			: P6215ma1.asp
'*  4. Program Name			: 금형점검계획 
'*  5. Program Desc			: 금형점검계획 등록 
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2005/01/19
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: Lee, SangHo
'* 10. Modifier (Last)		: Lee, SangHo
'* 11. Comment	
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									'☜: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim lgCheckall 
Dim lgCheckCase
Dim lgCheckDate
LocSvrDate = "<%=GetSvrDate%>"

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID			= "P6215mb1.asp"			'☆: List Production Order Header
Const BIZ_PGM_SAVE_ID			= "P6215mb2.asp"			'☆: Manage Production Order Header

Dim C_CHECK             '=  1
Dim C_FAC_CAST_CD		'=  2
Dim C_CAST_NM			'=  3
Dim C_SET_PLANT_CD		'=  4
Dim C_SET_PLANT_NM		'=  5
Dim C_CAR_KIND_CD		'=  6
Dim C_CAR_KIND_NM		'=  7
Dim C_MAKE_DT			'=  8
Dim C_INSP_PRID			'=  9
Dim C_CHECK_END_DT		'= 10
Dim C_FIN_CUR_ACCNT		'= 11
Dim C_FIN_AJ_DT			'= 12
Dim C_CUR_ACCNT			'= 13
Dim C_WORK_DT			'= 14
Dim C_WORK_DT_TEMP      '= 15

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgInvCloseDt	'재고마감일 
Dim lgCalType		'Calendar Type
Dim lgPlannedDate
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop						' Popup
Dim gSelframeFlg

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	'****************************
	'List Minor code(Order Type)
	'****************************
	<%
	Dim iData
    iData = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P3211' ")
	Response.write "Call SetCombo3(frm1.cboOrderType, """ &  iData & """) " & vbCrLf
	%>
	frm1.cboOrderType.value = "" 

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field

    Call InitSpreadSheet                                                    				'⊙: Setup the Spread sheet

	Call SetDefaultVal

    Call InitVariables																		'⊙: Initializes local global variables

    'Call InitComboBox()

    'Call InitSpreadComboBox()
    
    lgCheckall = 0
    lgCheckCase = 0
    lgCheckDate = 0
    
	frm1.btnSelectAll.disabled = True                                                       
	frm1.btnSelectCase.disabled = True
	frm1.btnSelectDate.disabled = True
	
    Call SetToolbar("1100000000001111")														'⊙: 버튼 툴바 제어 

	If parent.gPlant <> "" Then
		frm1.txtSetPlantCd.value = parent.gPlant
		frm1.txtSetPlantNm.value = parent.gPlantNm
		frm1.txtCarKind.focus()
		Set gActiveElement = document.activeElement
	Else
		frm1.txtSetPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If	
		
	
	Set gActiveElement = document.activeElement


End Sub

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

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtWork_dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtWork_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtWork_dt.Month = strMonth 
	frm1.txtWork_dt.Day = strDay	
	
	frm1.txtSpecial_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtSpecial_dt.Month = strMonth 
	frm1.txtSpecial_dt.Day = strDay
	
	'frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -10, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	'frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 20, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectCase.disabled = True
	frm1.btnSelectDate.disabled = True
	
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()    

    With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20030805", , Parent.gAllowDragDropSpread
		.vspdData.ReDraw = False
	
		.vspdData.MaxCols = C_WORK_DT_TEMP + 1
		.vspdData.MaxRows = 0
		Call GetSpreadColumnPos("A")
		 
		ggoSpread.SSSetCheck	C_CHECK,			"",	         		 2, , ,True ,-1
		ggoSpread.SSSetEdit		C_FAC_CAST_CD,		"금형코드",		18
		ggoSpread.SSSetEdit		C_CAST_NM,			"금형명",		40
		ggoSpread.SSSetEdit		C_SET_PLANT_CD,		"설치공장",		10
		ggoSpread.SSSetEdit		C_SET_PLANT_NM,		"설치공장",		20
		ggoSpread.SSSetEdit		C_CAR_KIND_CD,		"적용모델",		30
		ggoSpread.SSSetEdit		C_CAR_KIND_NM,		"적용모델",		30
		ggoSpread.SSSetDate 	C_MAKE_DT,			"제작일자",		11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_INSP_PRID,		"점검타수",		15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				
		ggoSpread.SSSetDate 	C_CHECK_END_DT,		"최종점검일",	11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_FIN_CUR_ACCNT,	"최종점검타수",	15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				
		ggoSpread.SSSetDate 	C_FIN_AJ_DT,		"현타수적용일",	11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_CUR_ACCNT,		"현재타수",		15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
		ggoSpread.SSSetDate 	C_WORK_DT,			"점검계획일",	11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_WORK_DT_TEMP,		"점검계획일",	11, 2, parent.gDateFormat

		.vspdData.ReDraw = True
		Call ggoSpread.SSSetColHidden(C_SET_PLANT_CD, C_SET_PLANT_CD, True)
		Call ggoSpread.SSSetColHidden(C_CAR_KIND_CD, C_CAR_KIND_CD, True)
		Call ggoSpread.SSSetColHidden(C_WORK_DT_TEMP, C_WORK_DT_TEMP, True)
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		ggoSpread.SSSetSplit2(2)

    End With

    Call SetSpreadLock()
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1
	ggoSpread.Source = .vspdData
	
	.vspdData.ReDraw = False
	ggoSpread.SpreadLock		C_FAC_CAST_CD,	-1, C_WORK_DT_TEMP		,-1
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
	.vspddata.ReDraw = True

    End With

	Call SetSpreadColor(1,1)
	
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 
    With frm1.vspdData 
    
    .Redraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SSSetProtected C_FAC_CAST_CD,         pvStartRow
    ggoSpread.SSSetProtected C_CAST_NM,				pvStartRow
    ggoSpread.SSSetProtected C_SET_PLANT_CD,		pvStartRow
    ggoSpread.SSSetProtected C_SET_PLANT_NM,		pvStartRow
    ggoSpread.SSSetProtected C_CAR_KIND_CD,			pvStartRow
    ggoSpread.SSSetProtected C_CAR_KIND_NM,			pvStartRow
    ggoSpread.SSSetProtected C_MAKE_DT,				pvStartRow
    ggoSpread.SSSetProtected C_INSP_PRID,			pvStartRow
    ggoSpread.SSSetProtected C_CHECK_END_DT,		pvStartRow
    ggoSpread.SSSetProtected C_FIN_CUR_ACCNT,		pvStartRow
    ggoSpread.SSSetProtected C_FIN_AJ_DT,			pvStartRow
    ggoSpread.SSSetProtected C_CUR_ACCNT,			pvStartRow

    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()	

    C_CHECK             =  1
	C_FAC_CAST_CD		=  2
	C_CAST_NM			=  3
	C_SET_PLANT_CD		=  4
	C_SET_PLANT_NM		=  5
	C_CAR_KIND_CD		=  6
	C_CAR_KIND_NM		=  7
	C_MAKE_DT			=  8
	C_INSP_PRID			=  9
	C_CHECK_END_DT		= 10
	C_FIN_CUR_ACCNT		= 11
	C_FIN_AJ_DT			= 12
	C_CUR_ACCNT			= 13
	C_WORK_DT			= 14
	C_WORK_DT_TEMP      = 15
End Sub
 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_CHECK				=	iCurColumnPos(1)
		C_FAC_CAST_CD		=	iCurColumnPos(2)
		C_CAST_NM			=	iCurColumnPos(3)
		C_SET_PLANT_CD		=	iCurColumnPos(4)
		C_SET_PLANT_NM		=	iCurColumnPos(5)
		C_CAR_KIND_CD		=	iCurColumnPos(6)
		C_CAR_KIND_NM		=	iCurColumnPos(7)
		C_MAKE_DT			=	iCurColumnPos(8)
		C_INSP_PRID			=	iCurColumnPos(9)
		C_CHECK_END_DT		=	iCurColumnPos(10)
		C_FIN_CUR_ACCNT		=	iCurColumnPos(11)
		C_FIN_AJ_DT			=	iCurColumnPos(12)
		C_CUR_ACCNT			=	iCurColumnPos(13)
		C_WORK_DT			=	iCurColumnPos(14)
		C_WORK_DT_TEMP      =   iCurColumnPos(15)
 	End Select
 
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

Function OpenSetPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSetPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSetPlantCd.focus
	
End Function


'------------------------------------------  OpenCast()  ------------------------------------------------
'	Name : OpenCast()
'	Description : Cast PopUp
'------------------------------------------------------------------------------------------------------------
Function OpenCast()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	IF frm1.txtSetPlantCd.value <> "" THEN
		Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtSetPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			frm1.txtSetPlantNm.value = ""
			IsOpenPop = False
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtSetPlantCd.focus
			Set gActiveElement = document.ActiveElement
			Exit Function
		ELSE
			frm1.txtSetPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtSetPlantNm.value = ""
		IsOpenPop = False
		Call DisplayMsgBox("971012", "X", "공장코드", "X")
		frm1.txtSetPlantCd.focus
		Set gActiveElement = document.ActiveElement
		Exit Function
	END IF 

		arrParam(0) = "금형코드"								' 팝업 명칭 
		arrParam(1) = "Y_CAST"											' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtCastCd.Value)		' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = "SET_PLANT = " & FilterVar(frm1.txtSetPlantCd.value, "''", "S")								' Where Condition
		arrParam(5) = "금형코드"								' TextBox 명칭 

    arrField(0) = "ED15" & parent.gcolsep & "CAST_CD"							' Field명(0)
    arrField(1) = "ED15" & parent.gcolsep & "CAST_NM"							' Field명(1)
    arrField(2) = "ED20" & parent.gcolsep & "(SELECT ITEM_GROUP_NM FROM B_ITEM_GROUP WHERE ITEM_GROUP_CD = CAR_KIND )"						' Field명(2)
    arrField(3) = "ED20" & parent.gcolsep & "(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = ITEM_CD_1 )"						' Field명(3)
    arrField(4) = "F3"   & parent.gcolsep & "EXT1_QTY"						' Field명(4)

    arrHeader(0) = "금형코드"					' Header명(0)
    arrHeader(1) = "금형코드명"					' Header명(1)
    arrHeader(2) = "모델명"						' Header명(2)
    arrHeader(3) = "품목명"						' Header명(3)
    arrHeader(4) = "차수"						' Header명(4)

	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCast(arrRet)
	End If

	Call SetFocusToDocument("M")
	frm1.txtCastCd.focus
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenCarKind()
'	Description : Condition CarKind PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCarKind()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "적용모델"						' 팝업 명칭 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCarKind.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "적용모델"						' TextBox 명칭 
	
    arrField(0) = "ITEM_GROUP_CD"						' Field명(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field명(1)
    
    arrHeader(0) = "적용모델"						' Header명(0)
    arrHeader(1) = "적용모델명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCarKind(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCarKind.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Function SetCast(byval arrRet)
	frm1.txtCastCd.Value    = arrRet(0)		
	frm1.txtCastNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetPlant()
'	Description : Condition SetPlant Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtSetPlantCd.Value    = arrRet(0)		
	frm1.txtSetPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind()
'	Description : Condition CarKind Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind(byval arrRet)
	frm1.txtCarKind.Value    = arrRet(0)		
	frm1.txtCarKindNm.Value  = arrRet(1)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function JumpOrderRun()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then					'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
    
   	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ReWorkFlag
	If frm1.vspdData.Text = "Y" Then
		Call DisplayMsgBox("189218", "x", "x", "x")
		Exit Function
	End If
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
   	frm1.vspdData.Col = C_ItemCode
	WriteCookie "txtItemCd", UCase(Trim(frm1.vspdData.Text))
	frm1.vspdData.Col = C_ItemName
	WriteCookie "txtItemNm", Trim(frm1.vspdData.Text)
	frm1.vspdData.Col = C_Specification
	WriteCookie "txtSpecification", Trim(frm1.vspdData.Text)
   	frm1.vspdData.Col = C_ProdtOrderNo
	WriteCookie "txtProdOrderNo", UCase(Trim(frm1.vspdData.Text))
   	frm1.vspdData.Col = C_PlanOrderNo
	WriteCookie "txtPlanOrderNo", UCase(Trim(frm1.vspdData.Text))
   	frm1.vspdData.Col = C_OrderQty
	WriteCookie "txtOrderQty", UCase(Trim(frm1.vspdData.Text))
	frm1.vspdData.Col = C_OrderUnit
	WriteCookie "txtOrderUnit", UCase(Trim(frm1.vspdData.Text))
   	frm1.vspdData.Col = C_PlanStartDt
	WriteCookie "txtPlanStartDt", UCase(Trim(frm1.vspdData.Text))
   	frm1.vspdData.Col = C_PlanEndDt
	WriteCookie "txtPlanEndDt", UCase(Trim(frm1.vspdData.Text))
	WriteCookie "txtInvCloseDt", lgInvCloseDt
	WriteCookie "txtPGMID", "P4112MA1"
	
	navigate BIZ_PGM_JUMPORDERRUN_ID
	
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

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtPlantCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtPlantCd_onChange()
	Call LookUpInvClsDt()
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row)

	Dim	DtPlanStartDt, DtPlanComptDt, DtInvCloseDt
	Dim strYear,strMonth,strDay
	Dim DtPlanStartDtDateFormat, DtPlanComptDtDateFormat
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
  	Else
  		Call SetPopupMenuItemInf("1001111111")         '화면별 설정 
  	End If
	
    With frm1.vspdData
		'----------------------
		'Column Split
		'----------------------
		gMouseClickStatus = "SPC"
	
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
 		End If
	
    End With

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

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub
 
'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  1
				.Col = Col
				intIndex = .Value
				.Col = C_BillFG
				.Value = intIndex
		End Select
	End With
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
         Exit Sub
	End If  
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		If Row < 1 Then Exit Sub

		Select Case Col

		    Case C_ItemPopup
				.Col = C_ItemCode
				.Row = Row
				Call OpenItemInfo2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_ItemCode,Row,"M","X","X")
				Set gActiveElement = document.activeElement
	    
		    Case C_TrackingNoPopup
				.Col = C_TrackingNo
				.Row = Row
				Call OpenTrackingInfo2(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_TrackingNo,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
		    Case C_RoutingNoPopup
				.Col = C_RoutingNo
				.Row = Row
				Call OpenRoutingNo(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_RoutingNo,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
		    Case C_SLCDPopup
				.Col = C_SLCD
				.Row = Row
				Call OpenSLCD(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_SLCD,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
		    Case C_OrderUnitPopup
				.Col = C_OrderUnit
				.Row = Row
				Call OpenUnit(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_OrderUnit,Row,"M","X","X")
				Set gActiveElement = document.activeElement

		End Select

    End With
    
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
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    'Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()

	Call InitData(1)
    
    With frm1.vspdData
		.ReDraw = False
		ggoSpread.SSSetProtected C_ItemCode,		-1, -1
		ggoSpread.SSSetProtected C_ItemPopup,		-1, -1
		ggoSpread.SSSetProtected C_ProdtOrderNo,	-1, -1
			
		If .MaxRows < 1 Then Exit Sub
		
		For LngRow = 1 To .MaxRows
			.Row = LngRow
			.Col = C_TrackingNo
			If .Text = "*" Or .Text = "" Then
				ggoSpread.SpreadLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
				ggoSpread.SSSetProtected C_TrackingNo, LngRow, LngRow
				ggoSpread.SSSetProtected C_TrackingNoPopup, LngRow, LngRow
			Else
			    ggoSpread.SpreadUnLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
				ggoSpread.SSSetRequired C_TrackingNo, LngRow, LngRow
			End If
			
		Next

		If lgIntFlgMode = parent.OPMD_CMODE Then

			.Row = 1
			.Col = C_OrderUnitMFG
			frm1.txtOrderUnitMFG.value = .Text
			.Col = C_MinMRPQty
			frm1.txtMinMRPQty.value = .Text
			.Col = C_FixedMRPQty
			frm1.txtFixedMRPQty.value = .Text
			.Col = C_MaxMRPQty
			frm1.txtMaxMRPQty.value = .Text
			.Col = C_RoundQty
			frm1.txtRoundQty.value = .Text
			.Col = C_ValidFromDT
			frm1.txtValidFromDT.Text = .Text
			.Col = C_ValidToDT
			frm1.txtValidToDT.Text = .Text
			.Col = C_OrderLtMFG
			frm1.txtOrderLtMFG.value = .Text
			.Col = C_ScrapRateMFG
			frm1.txtScrapRateMFG.value = .Text
			.Col = C_MPSMgr
			frm1.txtMPSMgr.value = .Text
			.Col = C_MRPMgr
			frm1.txtMRPMgr.value = .Text
			.Col = C_ProdMgr
			frm1.txtProdMgr.value = .Text
		End If
		.ReDraw = True
	End With
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

	FncQuery = False                                                        '⊙: Processing is NG
	    
	Err.Clear                                                               '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData										'⊙: Preset spreadsheet pointer 
	If ggoSpread.SSCheckChange = True Then									'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If


	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitVariables														'⊙: Initializes local global variables

	'-----------------------
	'Check condition area
	'-----------------------

	Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtSetPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	IF frm1.txtSetPlantCd.value <> "" THEN
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "공장코드", "X")
			frm1.txtSetPlantCd.focus
			frm1.txtSetPlantNm.value = ""
			Exit Function
		ELSE
			frm1.txtSetPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtSetPlantNm.value = ""
	END IF

	IF frm1.txtCarKind.value <> "" THEN
		Call  CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD = '" & frm1.txtCarKind.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "적용모델", "X")
			frm1.txtcarKind.focus
			Set gActiveElement = document.activeElement
			frm1.txtCarKindNm.value = ""
			Exit Function
		ELSE
			frm1.txtCarKindNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtCarKindNm.value = ""			
	END IF

	IF frm1.txtCastCd.value <> "" THEN
		Call  CommonQueryRs(" cast_nm "," y_cast "," SET_PLANT = '" & frm1.txtSetPlantCd.value & "' AND cast_cd = '" & frm1.txtCastCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "금형코드", "X")
			frm1.txtCastCd.focus
			Set gActiveElement = document.ActiveElement
			frm1.txtCastNm.value = ""
			Exit Function
		ELSE
			frm1.txtCastNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtCastNm.value = ""
	END IF
		
	If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
	Exit Function
	End If

	'-----------------------
	'Query function call area
	'-----------------------
		    
	If DbQuery = False Then Exit Function															'☜: Query db data
	    
	Call ggoOper.LockField(Document , "N")
	       
	FncQuery = True															'⊙: Processing is OK
   
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
    
    FncSave = False                                             '⊙: Processing is NG
    
    Err.Clear													'☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'☜: Save db data
    
    FncSave = True                                              '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
        
    
End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================

Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    Call initData(frm1.vspdData.ActiveRow)
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
	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData.MaxRows < 1 Then Exit Function

    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
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
    Call parent.FncExport(parent.C_MULTI)												
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()

    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
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
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'******************************************************************************************************%>
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	    
    Err.Clear

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    Dim strSetPlantCd
    Dim strCarKind
    Dim strCastCd
	
	If IsNull(frm1.txtSetPlantCd.value) Or Trim(frm1.txtSetPlantCd.value) = "" Then
		strSetPlantCd = "%"
	Else
		strSetPlantCd = Trim(frm1.txtSetPlantCd.value)
	End If

	If IsNull(frm1.txtCarKind.value) Or Trim(frm1.txtCarKind.value) = "" Then
		strCarKind = "%"
	Else
		strCarKind = Trim(frm1.txtCarKind.value)
	End If

	If IsNull(frm1.txtCastCd.value) Or Trim(frm1.txtCastCd.value) = "" Then
		strCastCd = "%"
	Else
		strCastCd = Trim(frm1.txtCastCd.value)
	End If

	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtSetPlantCd=" & strSetPlantCd
	strVal = strVal & "&txtCarKind=" & strCarKind
	strVal = strVal & "&txtCastCd=" & strCastCd
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

	Call DbQueryOk
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()															'☆: 조회 성공후 실행로직 

 	Dim lRow
 	Dim LngRow    

	
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field
    Call SetToolBar("1100100000001111")											'⊙: 버튼 툴바 제어 

	'frm1.vspdData.ReDraw = False
	'frm1.vspdData.ReDraw = True
	frm1.btnSelectAll.disabled = False
	frm1.btnSelectCase.disabled = False
	frm1.btnSelectDate.disabled = False
	lgCheckall = 0
	lgCheckCase = 0
	lgCheckDate = 0
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
   
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 성공적이지 아닐경우 
'========================================================================================
Function DbQueryNotOk()	

	Call SetToolBar("11001101001111")														'⊙: 버튼 툴바 제어 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE													'⊙: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

Dim lRow        
Dim strVal, strDel
Dim lColSep, lRowSep
Dim lGrpCnt  



lColSep = parent.gColSep
lRowSep = parent.gRowSep        
Err.Clear		
	
DbSave = False                                                   

With frm1.vspdData
   
	'-----------------------
	'Data manipulate area
	'-----------------------
		
	For lRow = 1 To .MaxRows
		.Row = lRow
		.Col = 0
		Select Case .Text
			Case ggoSpread.UpdateFlag
			
			.Col = C_Check
			If .Text = "1" Then	
				strVal = strVal & "U" & lColSep
				strVal = strVal & "20" & lColSep
				.Col = C_FAC_CAST_CD		: strVal = strVal & Trim(.Text) & lColSep
				.Col = C_WORK_DT	    : strVal = strVal & UNIConvDate(Trim(.Text)) & lColSep
				If Isnull(UNIConvDate(Trim(.Text))) or UNIConvDate(Trim(.Text)) = "" or UNIConvDate(Trim(.Text)) = "1900-01-01" Then
					IntRetCD = DisplayMsgBox("Y60020", "x", "x", "x")
					Call SetToolBar("1100100000001111")	
					Exit Function
				End If
				.Col = C_WORK_DT_TEMP   : strVal = strVal & UNIConvDate(Trim(.Text)) & lColSep
				strVal = strVal & "N" & lRowSep
				lGrpCnt = lGrpCnt + 1
		    End If
		End Select
	Next
End With
Call LayerShowHide(1)
frm1.txtMode.value        =  parent.UID_M0002
frm1.txtMaxRows.value     = lGrpCnt-1
frm1.txtSpread.value      = strVal

Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)						
	
DbSave = True                                                   
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

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

'========================================================================================
' Function Name : SelectAll
' Function Desc : 조회후, 전체데이터를 선택하여 준다.
'========================================================================================
Function SelectAll()
	
Dim IRowCount 
Dim IClnCount
Dim lWork_Dt_Temp 
Dim lWork_Dt


'lWork_Dt = UniConvDateToYYYYMMDD(frm1.txtWork_dt.text, Parent.gDateFormat, Parent.gComDateType)
lWork_Dt = frm1.txtWork_dt.text

ggoSpread.Source = frm1.vspdData
 
With frm1.vspdData   
	.ReDraw = False 
	IF lgCheckall = 0 Then 
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				If IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 1     
				Else
					.Row = IRowCount
					.Col = IClnCount
					.Text =ggoSpread.UpdateFlag
				End If
				.Col = C_WORK_DT
				.Text = lWork_Dt
			Next    
		Next
		lgCheckall = 1
		lgCheckCase = 0
		lgCheckDate = 0
	Else
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				if IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 0     
				Else
				End If
				.Col = C_WORK_DT_TEMP
				lWork_Dt_Temp = .Text
				.Col = C_WORK_DT
				.Text = lWork_Dt_Temp
			Next    
		Next
		lgCheckall = 0
		lgCheckCase = 0
		lgCheckDate = 0
	End If
	.ReDraw = True
End With

End Function		

'========================================================================================
' Function Name : SelectCase
' Function Desc : 조회후, 예상금형을 선택하여 준다.
'========================================================================================
Function SelectCase()
	
Dim IRowCount 
Dim IClnCount
Dim ldc_cur_accnt
Dim ldc_Fin_cur_accnt
Dim ldc_insp_prid
Dim lWork_Dt
Dim lWork_Dt_Temp
  
lWork_Dt = frm1.txtWork_dt.text

ggoSpread.Source = frm1.vspdData
 
With frm1.vspdData    
	If lgCheckCase = 0 Then 
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				If IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 0     
				Else
					.Row = IRowCount
					.Col = IClnCount
					.Text = ggoSpread.UpdateFlag
				End If
			Next    
		Next
		
		For IRowCount = 1 To .MaxRows
			.Row = IRowCount
			.Col = C_CUR_ACCNT
			ldc_cur_accnt     = .Value
			.Col = C_FIN_CUR_ACCNT 
			ldc_fin_cur_accnt = .Value
			.Col = C_INSP_PRID
			ldc_insp_prid = .Value
		    
			If ldc_cur_accnt - ldc_fin_cur_accnt  >=  ldc_insp_prid - (  ldc_insp_prid * 0.1) then
				.Col = C_CHECK 
				.Text = 1
				.Col = C_WORK_DT
				.Text = lWork_Dt
			End If	
		Next
		lgCheckCase = 1
		lgCheckAll = 0
		lgCheckDate = 0
	Else
   
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				If IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 0     
				Else
					.Row = IRowCount
					.Col = IClnCount
					.Text = ggoSpread.UpdateFlag
					.Col = C_WORK_DT_TEMP
					lWork_Dt_Temp = .Text
					.Col = C_WORK_DT
					.Text = lWork_Dt_Temp 
				End If
			Next    
		Next
		lgCheckCase = 0
		lgCheckAll = 0
		lgCheckDate = 0
	End If

End With

End Function		


'========================================================================================
' Function Name : SelectDate
' Function Desc : 조회후, 특정일자를 선택하여 준다.
'========================================================================================
Function SelectDate()
	
Dim IRowCount 
Dim IClnCount
Dim ldc_cur_accnt
Dim ldc_Fin_cur_accnt
Dim ldc_insp_prid
Dim lWork_Dt
Dim lSpecial_Dt
Dim lWork_Dt_Temp
Dim lSelect
Dim lCnt
Dim lWork_Dt_Cell

lCnt = 0

lWork_Dt = frm1.txtWork_dt.text
lSpecial_Dt = frm1.txtSpecial_dt.text

If IsNull(lSpecial_Dt) or lSpecial_Dt = "0000-00-00" or lSpecial_Dt = ""  then
	lSelect = Null
Else
	lSelect = lSpecial_Dt
End if

ggoSpread.Source = frm1.vspdData
 
With frm1.vspdData    
	If lgCheckDate = 0 Then 
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				If IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 0     
				Else
					.Row = IRowCount
					.Col = IClnCount
					.Text = ggoSpread.UpdateFlag
				End If
			Next    
		Next
		If IsNull(lselect) Then
			For IRowCount = 1 to .MaxRows
				.Row = IRowCount
				.Col = C_WORK_DT
				lSpecial_Dt = .Text
				If IsNull(lSpecial_Dt) or Trim(lSpecial_Dt) = "" Then
					.Col = C_CHECK 
					.Text = 1
					.Col = C_WORK_DT
					.Text = lWork_Dt 
					lCnt = lCnt + 1
				End If
			Next
		Elseif IsDate(lSelect) Then
			For IRowCount = 1 to .MaxRows
				.Row = IRowCount
				.Col = C_WORK_DT
				lSpecial_Dt = .Text
				If lSpecial_Dt = lSelect Then
					.Col = C_CHECK 
					.Text = 1
					.Col = C_WORK_DT
					.Text = lWork_Dt 
					lCnt = lCnt + 1
				End If
			Next
		Else
			IRowCount = 0
		End If
		

		
		If lCnt > 0 Then
			lgCheckDate = 1
			lgCheckAll = 0
			lgCheckCase = 0
		End If
	Else
		For IClnCount = 0 To C_CHECK
			For IRowCount = 1 To .MaxRows
				If IClnCount <> 0 Then   	     	 
					.Row = IRowCount 
					.Col = IClnCount	 
					.text = 0     
				Else
					.Row = IRowCount
					.Col = IClnCount
					.Text = ggoSpread.UpdateFlag
					.Col = C_WORK_DT_TEMP
					lWork_Dt_Temp = .Text
					.Col = C_WORK_DT
					.Text = lWork_Dt_Temp 
				End If
			Next    
		Next
		lgCheckCase = 0
		lgCheckAll = 0
		lgCheckDate = 0
	End If

End With

End Function		

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	Dim lWork_dt
	Dim lWork_dt_temp
	
	lWork_dt = frm1.txtWork_dt.Text

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_Check Then
			.Col = Col
			.Row = Row									
			IF .Text = 1 Then
				.Col = 0
				.Text = ggoSpread.UpdateFlag
				.Col = C_WORK_DT
				.Text = lWork_dt
				lgBlnFlgChgValue = True
			Elseif .Text = 0 Then
				.Col = 0
				.Text = ""
				.Col = C_WORK_DT_TEMP
				lWork_dt_temp = .Text
				.Col = C_WORK_DT
				.Text = lWork_dt_temp
				lgBlnFlgChgValue = False
			End if  		
		End If	
	End With
End Sub

'==========================================================================================
'   Event Name : 날짜 관련 더블클릭 이벤트 처리 모음 
'==========================================================================================

Sub txtWork_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtWork_dt.Action = 7
		frm1.txtWork_dt.focus
	End If
End Sub

Sub txtSpecial_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtSpecial_dt.Action = 7
		frm1.txtSpecial_dt.focus
	End If
End Sub
		
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형점검계획</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSetPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSetPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtSetPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>적용모델</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCarKind" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="적용모델"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCarKind" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCarKind()">&nbsp;<INPUT TYPE=TEXT NAME="txtCarKindNm" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>금형코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCastCd"  SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="금형코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCastCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCast()">&nbsp;<INPUT TYPE=TEXT NAME="txtCastNm" SIZE=20 tag="14" ALT="금형코드명"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>계획일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/p6215ma1_txtWork_dt_txtWork_dt.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>특정일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/p6215ma1_txtSpecial_dt_txtSpecial_dt.js'></script>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p6215ma1_A_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>				
					<TD WIDTH=10>&nbsp;</TD>
					<TD Align=left><BUTTON NAME="btnSelectAll" ONCLICK="vbscript:SelectAll()" CLASS="CLSMBTN">전체선택/취소</BUTTON>&nbsp
									<BUTTON NAME="btnSelectCase" ONCLICK="vbscript:SelectCase()" CLASS="CLSMBTN">예상금형선택/취소</BUTTON>&nbsp
									<BUTTON NAME="btnSelectDate" ONCLICK="vbscript:SelectDate()" CLASS="CLSMBTN">특정일자선택/취소</BUTTON></TD>
					<TD WIDTH=*></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
