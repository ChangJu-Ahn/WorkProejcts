<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name			: 원가 
'*  2. Function Name		:공정별수불조회 
'*  3. Program ID			: C4205MA1.asp
'*  4. Program Name			:공정별수불조회 
'*  5. Program Desc			: 
'*  6. Business ASP List	: +C4205Mb1.asp
'*						
'*  7. Modified date(First)	: 2005/09/16
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				: 
'* 12. History              : 
'*                          : 
'**********************************************************************************************-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID	= "c4205mb1.asp"			'☆:  비지니스 로직 ASP명 
Const BIZ_PGM_ID2	= "c4205mb2.asp"			'☆:  비지니스 로직 ASP명 


Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)					
								
'--spread A
Dim C_PlantCd		
Dim C_CCCd			
Dim C_CCNm		
Dim C_OrderNo		
Dim C_RoutNo
Dim C_CloseFlag
Dim C_Unit
Dim C_WcCd
Dim C_WcNm
Dim C_OprNo
Dim C_InsideFlg
Dim C_MilestoneFlg
Dim C_WipQty
Dim C_PriorOprQty
Dim C_NextOprQty
Dim C_LastWipQty
Dim C_BAS_BAD_Qty
Dim C_THIS_BadQty
Dim C_REWORKED_BAD_QTY
Dim C_BAL_BAD_QTY
Dim C_ProdRate		


'--spread B
Dim C_WcCd2
Dim C_WcNm2
Dim C_OprNo2
Dim C_AcctNm2
Dim C_ItemCd2
Dim C_ItemNm2
Dim C_Unit2
Dim C_WipQty2
Dim C_WipAmt2
Dim C_WipPrice2


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey
Dim lgIntPrevKey
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim IsOpenPop						' Popup


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

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
	lgBlnFlgChgValue = False
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgSortKey = 1

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(byVal pvSpd)	
    
    Call InitSpreadPosVariables(pvSpd)
    Call AppendNumberPlace("6","3","0")
    
    If pvSpd = "" or pvSpd ="A" Then 
    
		With frm1
		       
		ggoSpread.Source = .vspdData
		'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021106", , ""
     
		.vspdData.ReDraw = False
    
		.vspdData.MaxCols = C_ProdRate + 1
		.vspdData.MaxRows = 0
	
		Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetEdit		C_PlantCd,				"공장", 10,,,7
		ggoSpread.SSSetEdit		C_CCCd,				"작업지시C/C", 10   
		ggoSpread.SSSetEdit		C_CCNM,				"작업지시C/C명", 20
		ggoSpread.SSSetEdit		C_OrderNo,			"제조번호", 18
		ggoSpread.SSSetEdit		C_RoutNo,				"라우팅번호", 10   
		ggoSpread.SSSetEdit		C_CloseFlag,				"마감여부", 10
		ggoSpread.SSSetEdit		C_Unit,				"재고단위", 10   
		ggoSpread.SSSetEdit		C_WcCd,				"공정", 10
		ggoSpread.SSSetEdit		C_WcNm,				"공정명", 20
		ggoSpread.SSSetEdit		C_OprNo,				"공순", 10
		ggoSpread.SSSetEdit		C_InsideFlg,				"사내공정여부", 10  
		ggoSpread.SSSetEdit		C_MilestoneFlg,				"Milestone", 10
    
		ggoSpread.SSSetFloat		C_WipQty,				"기초재공수량",		15,parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_PriorOprQty,				"전공정대체(착수)수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_NextOprQty,				"차공정(완성)대체수량",	15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_LastWipQty,				"기말재공수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_BAS_BAD_Qty,				"기초불량수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_THIS_BadQty,				"당월불량수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_REWORKED_BAD_QTY,				"재작업불량수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_BAL_BAD_QTY,				"기말불량수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
    

		ggoSpread.SSSetFloat		C_ProdRate,			"완성품환산율(%)", 15,"6",	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,,	"Z"  ,"0","100"
   
 		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols ,.vspdData.MaxCols , True)
			
		
		.vspdData.ReDraw =True
	
		End With
   
		Call SetSpreadLock("A")
	End If
	
	 If pvSpd = "" or pvSpd ="B" Then 
    
		With frm1
		       
		ggoSpread.Source = .vspdData2
		'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		ggoSpread.Spreadinit "V20021106", , ""
    
		'Call AppendNumberPlace("6","3","0")
		  
    
		.vspdData2.ReDraw = False
    
		.vspdData2.MaxCols = C_WipPrice2 + 1
		.vspdData2.MaxRows = 0
	
		Call GetSpreadColumnPos("B")
	
		ggoSpread.SSSetEdit		C_WcCd2,				"공정", 10
		ggoSpread.SSSetEdit		C_WcNm2,				"공정명", 20    
		ggoSpread.SSSetEdit		C_OprNo2,				"공순", 8
		ggoSpread.SSSetEdit		C_AcctNm2,			"품목계정", 18
		ggoSpread.SSSetEdit		C_ItemCd2,				"투입품목", 10   
		ggoSpread.SSSetEdit		C_ItemNm2,				"투입품목명", 30
		ggoSpread.SSSetEdit		C_Unit2,				"재고단위", 10,,,7    
		ggoSpread.SSSetFloat		C_WipQty2,				"투입수량",		15,		parent.ggQtyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_WipAmt2,				"투입금액",		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_WipPrice2,				"투입단가",   15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		
 		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols ,.vspdData2.MaxCols , True)
			
		'ggoSpread.SSSetSplit2(2) 
		.vspdData2.ReDraw =True
	
		End With
   
		Call SetSpreadLock("B")
	End If
    
End Sub
'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : ReInitSpreadSheet
' Function Desc : This method re-initializes spread sheet column property
'========================================================================================
Sub ReInitSpreadSheet()
	
	Dim ret, iRowSpan,i
	
	With frm1.vspdData

		.Col = .MaxCols
		.ColHidden = True
		
		
		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1			
	
	End With
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock(byVal pvSpd)
	If pvSpd="A" Then
		ggoSpread.Source = frm1.vspdData    
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
	If pvSpd="B" Then
		ggoSpread.Source = frm1.vspdData2    
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(byVal pvSpd)
	If pvSpd="" or pvSpd="A" Then
		C_PlantCd	      =1
		C_CCCd		=2
		C_CCNm		=3
		C_OrderNo	=4	
		C_RoutNo      =5
		C_CloseFlag   =6
		C_Unit           =7
		C_WcCd        =8
		C_WcNm       =9
		C_OprNo        =10
		C_InsideFlg     =11
		C_MilestoneFlg  =12
		C_WipQty         =13
		C_PriorOprQty   =14
		C_NextOprQty    =15		
		C_LastWipQty     =16		
		C_BAS_BAD_Qty       =17
		C_THIS_BadQty          =18
		C_REWORKED_BAD_QTY = 19
		C_BAL_BAD_QTY = 20
		C_ProdRate	     =21

	'C_GoodsQty       =15

	
	End If
	
	If pvSpd="" or pvSpd="B" Then	
		C_WcCd2       =1
		C_WcNm2      =2
		C_OprNo2      =3
		C_AcctNm2    =4
		C_ItemCd2     =5
		C_ItemNm2    =6
		C_Unit2          =7
		C_WipQty2     =8
		C_WipAmt2    =9
		C_WipPrice2	 =10
	End If

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
		C_PlantCd	     = iCurColumnPos(1)	
		C_CCCd		 = iCurColumnPos(2)	
		C_CCNm		 = iCurColumnPos(3)
		C_OrderNo	 = iCurColumnPos(4)	
		C_RoutNo       = iCurColumnPos(5)
		C_CloseFlag    = iCurColumnPos(6)
		C_Unit            = iCurColumnPos(7)
		C_WcCd         = iCurColumnPos(8)
		C_WcNm         = iCurColumnPos(9)
		C_OprNo         = iCurColumnPos(10)
		C_InsideFlg     = iCurColumnPos(11)
		C_MilestoneFlg  = iCurColumnPos(12)
		C_WipQty         =iCurColumnPos(13)
		C_PriorOprQty   =iCurColumnPos(14)
		C_NextOprQty    =iCurColumnPos(15)	
		C_LastWipQty     =iCurColumnPos(16)	
		C_BAS_BAD_Qty       =iCurColumnPos(17)
		C_THIS_BadQty          =iCurColumnPos(18)
		C_REWORKED_BAD_QTY = iCurColumnPos(19)
		C_BAL_BAD_QTY = iCurColumnPos(20)
		C_ProdRate	     =iCurColumnPos(21)
		
	
	Case "B"
 		ggoSpread.Source = frm1.vspdData2  		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos) 		
 		
 		C_WcCd2   = iCurColumnPos(1)
		C_WcNm2  = iCurColumnPos(2)
		C_OprNo2  = iCurColumnPos(3)
		C_AcctNm2  = iCurColumnPos(4)
		C_ItemCd2  = iCurColumnPos(5)
		C_ItemNm2  = iCurColumnPos(6)
		C_Unit2       = iCurColumnPos(7)
		C_WipQty2  = iCurColumnPos(8)
		C_WipAmt2  = iCurColumnPos(9)
		C_WipPrice2	= iCurColumnPos(10)
 		 		
 	End Select 
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
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
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"						' Header명(0)
    arrHeader(1) = "공장명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "OP"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenCCCd()  -------------------------------------------------
'	Name : OpenCCCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCCCd()
	Dim arrRet
	Dim strWhere
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	strWhere =""
	IsOpenPop = True
	If frm1.txtPlantCd.value<>"" Then
	strWhere= " plant_cd=" & FilterVar(Trim(frm1.txtPlantCd.Value),"''","S")	
	end If

	
	arrParam(0) = "작업지시C/C"						' 팝업 명칭 
	arrParam(1) = "B_COST_CENTER "						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCCCd.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =strWhere							' Where Condition
	arrParam(5) = "C/C"							' TextBox 명칭 
	
    arrField(0) ="ED10" & Parent.gColSep &  "COST_CD"					' Field명(0)
    arrField(1) = "ED31" & Parent.gColSep & "COST_NM"					' Field명(1)
    
    
    arrHeader(0) = "C/C"						' Header명(0)
    arrHeader(1) = "C/C명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCCCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCCCd.focus
	
End Function



'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
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

'------------------------------------------  SetCCCd()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCCCd(byval arrRet)
	frm1.txtCCCd.Value    = arrRet(0)		
	frm1.txtCCNm.Value   = arrRet(1)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet("")                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------	
	Call SetToolBar("11000000000111")											'⊙: 버튼 툴바 제어	
	frm1.txtYYYYMM.Text=LocSvrDate
	
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtCCCd.focus 		
	Else
		frm1.txtPlantCd.focus 		
	End If
	frm1.txtYYYYMM.Text=LocSvrDate		
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
				
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

'=========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
    
End Sub
'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정		
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData        
    
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
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================

Sub vspdData2_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
	End If

End Sub
'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = 1
           .vspdData2.MaxRows = 0
        End With
        frm1.vspddata.Col = 0

		Call DbDtlQuery(NewRow)
    End If
End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    'If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
     '        Exit Sub
	'End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgIntPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If LayerShowHide(1) = False Then Exit Sub
			If DbDtlQuery = False Then	
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
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
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
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
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
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.id)    
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtYYYYMM_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtYYYYMM.Action = 7
        Call SetFocusToDocument("P")
		Frm1.txtYYYYMM.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtYYYYMM_KeyDown
'   Event Desc : 
'=======================================================================================================
Sub  txtYYYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
'********************************************************************************************************* %>
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData										'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then									'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    

	IF ChkKeyField()=False Then Exit Function 
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call InitVariables		

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If															'☜: Query db data
       
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
   On Error Resume Next
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status      
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
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, i
    Dim iIntIndex
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                  '☜: Clear error status
     

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows, lDelRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												<%'☜: 화면 유형 %>
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function


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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function  FncSplitColumn()

'    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
 '      Exit Sub
  '  End If
'
 '   ggoSpread.Source = gActiveSpdSheet
  '  ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
	If gMouseClickStatus = "SPCRP" Then
		iColumnLimit = 3
       
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow

		If ACol > iColumnLimit Then
		   iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
		   Exit Function  
		End If   
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
		ggoSpread.Source = Frm1.vspdData
    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
    
		Frm1.vspdData.Action = 0    
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
	'----------------------------------------
	' Spread가 두개일 경우 2번째 Spread
	'----------------------------------------
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = 4
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   

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
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    Dim strGubun 
    
    If frm1.rdoBatch.checked Then
		strGubun = frm1.rdoBatch.value
	Else
		strGubun = frm1.rdoCost.value
	End If
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(frm1.hYYYYMM.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCCCd=" & Trim(frm1.hCCCd.value)				'☆: 조회 조건 데이타		
		strVal = strVal & "&txtGubun=" & Trim(strGubun)				'☆: 조회 조건 데이타	
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtYYYYMM=" & Trim(frm1.txtYYYYMM.text)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCCCd=" & Trim(frm1.txtCCCd.value)				'☆: 조회 조건 데이타	
		strVal = strVal & "&txtGubun=" & Trim(strGubun)				'☆: 조회 조건 데이타	
	End If

    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)															'☆: 조회 성공후 실행로직	
	
	If LngMaxRow <1 Then
		Exit Function 
	End If	
	Call ReInitSpreadSheet()
	lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode
    
    'Call vspdData_Click(1,1)
    Call vspdData_ScriptLeaveCell( 0,  0,  1,  1, "")
    
    frm1.vspdData.focus
																						'⊙: This function lock the suitable field
	Call ggoOper.LockField(Document, "Q")															'⊙: This function lock the suitable field 
End Function
'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(byval iRow) 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOrderNo
Dim strOprNo
Dim strGubun
    
	boolExist = False
    With frm1

		.vspdData2.MaxRows = 0

	    .vspdData.Row = iRow'.vspdData.ActiveRow
	    .vspdData.Col = C_OrderNo
	    strOrderNo = .vspdData.Text
	    .vspdData.Col = C_OprNo
	    strOprNo = .vspdData.Text    
	    
		If frm1.rdoBatch.checked Then
			strGubun = frm1.rdoBatch.value
		Else
			strGubun = frm1.rdoCost.value
		End If
		DbDtlQuery = False   
    
		.vspdData.Row =iRow' .vspdData.ActiveRow

		If LayerShowHide(1) = False Then Exit Function

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtYYYYMM=" & Trim(frm1.txtYYYYMM.Text)
			strVal = strVal & "&txtProdOrderNo=" & Trim(strOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtGubun=" & Trim(strGubun)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtSPId=" & Trim(frm1.hSpId.value)				'☆: 조회 조건 데이타		
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtYYYYMM=" & Trim(frm1.txtYYYYMM.Text)
			strVal = strVal & "&txtProdOrderNo=" & Trim(strOrderNo)
			strVal = strVal & "&txtOprNo=" & Trim(strOprNo)
			strVal = strVal & "&txtGubun=" & Trim(strGubun)				'☆: 조회 조건 데이타	
			strVal = strVal & "&txtSPId=" & Trim(frm1.hSpId.value)				'☆: 조회 조건 데이타		
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
     ggoSpread.Source = frm1.vspdData2   
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	
End Function
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

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

End Function


'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check plant
	If Trim(frm1.txtPlantCd.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlantCd.alt,"X")
			frm1.txtPlantCd.focus 
			frm1.txtPlantnm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNM.value = strDataNm(0)
	End If
'check CC cd	
	If Trim(frm1.txtCCCd.value) <> "" Then
		strWhere = " COST_Cd = " & FilterVar(frm1.txtCCCd.value, "''", "S") & " "		
		iF trim(frm1.txtPlantCd.value)<>"" Then 
			strWhere = strWhere & " and PLANT_CD=" & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "		
		End If
		Call CommonQueryRs(" COST_Nm ","	 B_COST_CENTER  ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCCCd.alt,"X")
			frm1.txtCCCd.focus 
			frm1.txtCCNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCCNm.value = strDataNm(0)
	End If
'check prod order no	
	If Trim(frm1.txtProdOrderNo.value) <> "" Then
		If  trim(frm1.txtPlantCd.value)="" Then 
			Call DisplayMsgBox("970000","X",frm1.txtPlantCd.alt,"X")
			Exit function 
		End If
		
		strFrom = " p_production_order_header a, b_item b,b_storage_location c,b_item_by_plant d "
		strWhere = " a.prodt_order_no = " & FilterVar(frm1.txtProdOrderNo.value, "''", "S") & " "		
		strWhere =strWhere & " and a.item_cd = b.item_cd and	a.plant_cd = d.plant_cd and	a.item_cd = d.item_cd and	a.sl_cd = c.sl_cd"	
		strWhere =strWhere & " AND a.PLANT_CD = " & FilterVar(trim(frm1.txtPlantCd.value), "''", "S") 
		strWhere =strWhere & " and 	a.order_status in (  'OP', 'RL', 'RL' ) "
		
		Call CommonQueryRs(" a.prodt_order_no ",strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtProdOrderNo.alt,"X")
			frm1.txtProdOrderNo.focus 
			frm1.txtProdOrderNo.value = ""
			ChkKeyField = False
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtProdOrderNo.value = strDataNm(0)
	End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>작업년월</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME tag="12" ALT="작업년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작업지시C/C</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCCd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCCCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtCCNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>									
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>구분</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO id="rdoCost"   NAME="rdoGubun" tag="11" CLASS="RADIO" value="A" checked >원가결산집계 &nbsp; <INPUT TYPE=RADIO id="rdoBatch" NAME="rdoGubun" CLASS="RADIO" tag="11" value="B">Batch집계
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="*">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 ID = "B" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=bizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hCCCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24"><INPUT TYPE=HIDDEN NAME="hSpId" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
