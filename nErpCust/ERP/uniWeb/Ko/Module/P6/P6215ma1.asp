<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: �������� 
'*  2. Function Name		: �������˰�ȹ 
'*  3. Program ID			: P6215ma1.asp
'*  4. Program Name			: �������˰�ȹ 
'*  5. Program Desc			: �������˰�ȹ ��� 
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
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									'��: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim lgCheckall 
Dim lgCheckCase
Dim lgCheckDate
LocSvrDate = "<%=GetSvrDate%>"

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID			= "P6215mb1.asp"			'��: List Production Order Header
Const BIZ_PGM_SAVE_ID			= "P6215mb2.asp"			'��: Manage Production Order Header

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

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgInvCloseDt	'������� 
Dim lgCalType		'Calendar Type
Dim lgPlannedDate
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

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
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                          			'��: Lock  Suitable  Field

    Call InitSpreadSheet                                                    				'��: Setup the Spread sheet

	Call SetDefaultVal

    Call InitVariables																		'��: Initializes local global variables

    'Call InitComboBox()

    'Call InitSpreadComboBox()
    
    lgCheckall = 0
    lgCheckCase = 0
    lgCheckDate = 0
    
	frm1.btnSelectAll.disabled = True                                                       
	frm1.btnSelectCase.disabled = True
	frm1.btnSelectDate.disabled = True
	
    Call SetToolbar("1100000000001111")														'��: ��ư ���� ���� 

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
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()

	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtWork_dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtWork_dt.Year = strYear 		 '����� default value setting
	frm1.txtWork_dt.Month = strMonth 
	frm1.txtWork_dt.Day = strDay	
	
	frm1.txtSpecial_dt.Year = strYear 		 '����� default value setting
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
		ggoSpread.SSSetEdit		C_FAC_CAST_CD,		"�����ڵ�",		18
		ggoSpread.SSSetEdit		C_CAST_NM,			"������",		40
		ggoSpread.SSSetEdit		C_SET_PLANT_CD,		"��ġ����",		10
		ggoSpread.SSSetEdit		C_SET_PLANT_NM,		"��ġ����",		20
		ggoSpread.SSSetEdit		C_CAR_KIND_CD,		"�����",		30
		ggoSpread.SSSetEdit		C_CAR_KIND_NM,		"�����",		30
		ggoSpread.SSSetDate 	C_MAKE_DT,			"��������",		11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_INSP_PRID,		"����Ÿ��",		15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				
		ggoSpread.SSSetDate 	C_CHECK_END_DT,		"����������",	11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_FIN_CUR_ACCNT,	"��������Ÿ��",	15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				
		ggoSpread.SSSetDate 	C_FIN_AJ_DT,		"��Ÿ��������",	11, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_CUR_ACCNT,		"����Ÿ��",		15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
		ggoSpread.SSSetDate 	C_WORK_DT,			"���˰�ȹ��",	11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_WORK_DT_TEMP,		"���˰�ȹ��",	11, 2, parent.gDateFormat

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

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'*********************************************************************************************************

Function OpenSetPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtSetPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
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
			Call DisplayMsgBox("971012", "X", "�����ڵ�", "X")
			frm1.txtSetPlantCd.focus
			Set gActiveElement = document.ActiveElement
			Exit Function
		ELSE
			frm1.txtSetPlantNm.value = left(lgF0, len(lgF0) -1)
		END IF
	ELSE
		frm1.txtSetPlantNm.value = ""
		IsOpenPop = False
		Call DisplayMsgBox("971012", "X", "�����ڵ�", "X")
		frm1.txtSetPlantCd.focus
		Set gActiveElement = document.ActiveElement
		Exit Function
	END IF 

		arrParam(0) = "�����ڵ�"								' �˾� ��Ī 
		arrParam(1) = "Y_CAST"											' TABLE ��Ī 
		arrParam(2) = Trim(frm1.txtCastCd.Value)		' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = "SET_PLANT = " & FilterVar(frm1.txtSetPlantCd.value, "''", "S")								' Where Condition
		arrParam(5) = "�����ڵ�"								' TextBox ��Ī 

    arrField(0) = "ED15" & parent.gcolsep & "CAST_CD"							' Field��(0)
    arrField(1) = "ED15" & parent.gcolsep & "CAST_NM"							' Field��(1)
    arrField(2) = "ED20" & parent.gcolsep & "(SELECT ITEM_GROUP_NM FROM B_ITEM_GROUP WHERE ITEM_GROUP_CD = CAR_KIND )"						' Field��(2)
    arrField(3) = "ED20" & parent.gcolsep & "(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = ITEM_CD_1 )"						' Field��(3)
    arrField(4) = "F3"   & parent.gcolsep & "EXT1_QTY"						' Field��(4)

    arrHeader(0) = "�����ڵ�"					' Header��(0)
    arrHeader(1) = "�����ڵ��"					' Header��(1)
    arrHeader(2) = "�𵨸�"						' Header��(2)
    arrHeader(3) = "ǰ���"						' Header��(3)
    arrHeader(4) = "����"						' Header��(4)

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

	arrParam(0) = "�����"						' �˾� ��Ī 
	arrParam(1) = "B_ITEM_GROUP"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCarKind.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "�����"						' TextBox ��Ī 
	
    arrField(0) = "ITEM_GROUP_CD"						' Field��(0)
    arrField(1) = "ITEM_GROUP_NM"						' Field��(1)
    
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "����𵨸�"						' Header��(1)
    
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
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCast()
'	Description : Cast POPUP���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetCast(byval arrRet)
	frm1.txtCastCd.Value    = arrRet(0)		
	frm1.txtCastNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetPlant()
'	Description : Condition SetPlant Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtSetPlantCd.Value    = arrRet(0)		
	frm1.txtSetPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetConPlant()  ------------------------------------------------
'	Name : SetCarKind()
'	Description : Condition CarKind Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetCarKind(byval arrRet)
	frm1.txtCarKind.Value    = arrRet(0)		
	frm1.txtCarKindNm.Value  = arrRet(1)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function JumpOrderRun()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData                        '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then					'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '��: Display Message(There is no changed data.)
        Exit Function
    End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
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
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
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
  		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
  	Else
  		Call SetPopupMenuItemInf("1001111111")         'ȭ�麰 ���� 
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
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
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
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
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
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------
	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
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
	
	If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
         Exit Sub
	End If  
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
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
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
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
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery()
 
	Dim IntRetCD 

	FncQuery = False                                                        '��: Processing is NG
	    
	Err.Clear                                                               '��: Protect system from crashing

	ggoSpread.Source = frm1.vspdData										'��: Preset spreadsheet pointer 
	If ggoSpread.SSCheckChange = True Then									'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If


	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	Call InitVariables														'��: Initializes local global variables

	'-----------------------
	'Check condition area
	'-----------------------

	Call  CommonQueryRs(" plant_nm "," b_plant "," plant_cd = '" & frm1.txtSetPlantCd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	IF frm1.txtSetPlantCd.value <> "" THEN
		IF IsNull(lgF0) OR Trim(lgF0) = "" THEN
			Call DisplayMsgBox("971012", "X", "�����ڵ�", "X")
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
			Call DisplayMsgBox("971012", "X", "�����", "X")
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
			Call DisplayMsgBox("971012", "X", "�����ڵ�", "X")
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
		
	If Not chkfield(Document, "1") Then								'��: This function check indispensable field
	Exit Function
	End If

	'-----------------------
	'Query function call area
	'-----------------------
		    
	If DbQuery = False Then Exit Function															'��: Query db data
	    
	Call ggoOper.LockField(Document , "N")
	       
	FncQuery = True															'��: Processing is OK
   
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
    
    FncSave = False                                             '��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'��: Save db data
    
    FncSave = True                                              '��: Processing is OK
    
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
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
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

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
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
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()															'��: ��ȸ ������ ������� 

 	Dim lRow
 	Dim LngRow    

	
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    Call SetToolBar("1100100000001111")											'��: ��ư ���� ���� 

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
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
   
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� ���������� �ƴҰ�� 
'========================================================================================
Function DbQueryNotOk()	

	Call SetToolBar("11001101001111")														'��: ��ư ���� ���� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE													'��: Indicates that current mode is Update mode

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
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'========================================================================================
' Function Name : SelectAll
' Function Desc : ��ȸ��, ��ü�����͸� �����Ͽ� �ش�.
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
' Function Desc : ��ȸ��, ��������� �����Ͽ� �ش�.
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
' Function Desc : ��ȸ��, Ư�����ڸ� �����Ͽ� �ش�.
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
'   Event Name : ��¥ ���� ����Ŭ�� �̺�Ʈ ó�� ���� 
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
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������˰�ȹ</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSetPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSetPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSetPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtSetPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCarKind" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCarKind" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCarKind()">&nbsp;<INPUT TYPE=TEXT NAME="txtCarKindNm" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCastCd"  SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCastCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCast()">&nbsp;<INPUT TYPE=TEXT NAME="txtCastNm" SIZE=20 tag="14" ALT="�����ڵ��"></TD>
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
									<TD CLASS="TD5" NOWRAP>��ȹ����</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/p6215ma1_txtWork_dt_txtWork_dt.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>Ư������</TD>
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
					<TD Align=left><BUTTON NAME="btnSelectAll" ONCLICK="vbscript:SelectAll()" CLASS="CLSMBTN">��ü����/���</BUTTON>&nbsp
									<BUTTON NAME="btnSelectCase" ONCLICK="vbscript:SelectCase()" CLASS="CLSMBTN">�����������/���</BUTTON>&nbsp
									<BUTTON NAME="btnSelectDate" ONCLICK="vbscript:SelectDate()" CLASS="CLSMBTN">Ư�����ڼ���/���</BUTTON></TD>
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
