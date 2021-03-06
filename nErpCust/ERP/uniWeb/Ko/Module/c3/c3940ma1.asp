<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : C3940MA1.asp
'*  4. Program Name         : 직과가공비배부내역조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "c3940mb1.asp"							'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_QRY2_ID	= "c3940mb2.asp"							'☆: 비지니스 로직 ASP명 

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

' Grid 1(vspdData1) 
Dim C_CostCd
Dim C_CostNm
Dim C_AcctCd
Dim C_AcctNm
Dim C_CtrlCd
Dim C_CtrlNm
Dim C_CtrlVal
Dim C_CtrlValNm
Dim C_AllocAmt

' Grid 2(vspdData2) 
Dim C_OrderNo
Dim C_Seq
Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_DstbFctr
Dim C_DstbFlag
Dim C_DstbBasis
Dim C_DstbAmt
Dim C_AdjRate

dim	strYYYYMM  

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size를 조사할 변수 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

Dim strDate
Dim iDBSYSDate
Dim lgCloseFlgMode

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    frm1.hSpid.value = ""
    lgCloseFlgMode	 = "O"			'C : Spid 삭제 O : 초기화 
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	
	frm1.txtYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)


	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_AllocAmt + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_CostCd,		"코스트센터"	,10 
			ggoSpread.SSSetEdit 	C_CostNm,       "코스트센터명"	,20
			ggoSpread.SSSetEdit 	C_AcctCd,       "계정"			,15 
			ggoSpread.SSSetEdit 	C_AcctNm,		"계정명"		,20
			ggoSpread.SSSetEdit 	C_CtrlCd,		"관리항목"		,8 
			ggoSpread.SSSetEdit 	C_CtrlNm,		"관리항목명"	,10
			ggoSpread.SSSetEdit 	C_CtrlVal,		"관리항목 VALUE",15
			ggoSpread.SSSetEdit 	C_CtrlValNm,	"관리항목 VALUE명",20
			ggoSpread.SSSetFloat 	C_AllocAmt,		"배부대상금액"	,15,parent.ggAmtofMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			
			
			Call ggoSpread.MakePairsColumn(C_CostCd, C_CostNm )
			Call ggoSpread.MakePairsColumn(C_AcctCd, C_AcctNm )
			Call ggoSpread.MakePairsColumn(C_CtrlCd, C_CtrlNm )
			Call ggoSpread.MakePairsColumn(C_CtrlVal, C_CtrlValNm )
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(4)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true    
    
		End With
	
    End If
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData2 
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021225", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_AdjRate + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_OrderNo,		"오더번호"		,15
			ggoSpread.SSSetEdit		C_Seq,			"SEQ"			,8
			ggoSpread.SSSetEdit		C_PlantCd,		"공장"			,10
			ggoSpread.SSSetEdit		C_PlantNm,		"공장명"		,20
			ggoSpread.SSSetEdit		C_ItemCd,		"품목"			,13
			ggoSpread.SSSetEdit 	C_ItemNm,		"품목명"		,20
			ggoSpread.SSSetEdit		C_DstbFctr,     "배부요소"		,15
			ggoSpread.SSSetEdit		C_DstbFlag,		"배부구분"		,20
			ggoSpread.SSSetFloat	C_DstbBasis,	"배부기준Data"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_DstbAmt,		"배부금액"		,15,parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_AdjRate,		"가중치"		,10,parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			
			Call ggoSpread.MakePairsColumn(C_OrderNo, C_Seq )
			Call ggoSpread.MakePairsColumn(C_PlantCd, C_PlantNm )
			Call ggoSpread.MakePairsColumn(C_ItemCd, C_ItemNm )
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true    
    
		End With
    End If
    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	With frm1
		If pvSpdNo = "A" Then
		   ggoSpread.Source = frm1.vspdData1
		   ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "B" Then 
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
		   
    End With
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub



'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Order Header
		C_CostCd			= 1
		C_CostNm			= 2
		C_AcctCd			= 3
		C_AcctNm			= 4
		C_CtrlCd			= 5
		C_CtrlNm			= 6
		C_CtrlVal			= 7
		C_CtrlValNm			= 8
		C_AllocAmt			= 9
	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Result
		C_OrderNo			= 1
		C_Seq				= 2
		C_PlantCd			= 3
		C_PlantNm			= 4
		C_ItemCd			= 5
		C_ItemNm			= 6
		C_DstbFctr			= 7
		C_DstbFlag			= 8
		C_DstbBasis			= 9
		C_DstbAmt			= 10
		C_AdjRate			= 11
	End If	

End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData1
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	
			C_CostCd			= iCurColumnPos(1)
			C_CostNm			= iCurColumnPos(2)
			C_AcctCd			= iCurColumnPos(3)
			C_AcctNm			= iCurColumnPos(4)
			C_CtrlCd			= iCurColumnPos(5)
			C_CtrlNm			= iCurColumnPos(6)
			C_CtrlVal			= iCurColumnPos(7)
			C_CtrlValNm			= iCurColumnPos(8)
	
		Case "B"
		
			ggoSpread.Source = frm1.vspdData2
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_OrderNo			= iCurColumnPos(1)
			C_Seq				= iCurColumnPos(2)
			C_PlantCd			= iCurColumnPos(3)
			C_PlantNm			= iCurColumnPos(4)
			C_ItemCd			= iCurColumnPos(5)
			C_ItemNm			= iCurColumnPos(6)
			C_DstbFctr			= iCurColumnPos(7)
			C_DstbFlag			= iCurColumnPos(8)
			C_DstbBasis			= iCurColumnPos(9)
			C_DstbAmt			= iCurColumnPos(10)
			C_AdjRate			= iCurColumnPos(11)

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
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strYYYYMM1,strYear,strMonth,strDay
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMM1 = strYear & strMonth
	
	select case iWhere
		case 1
			arrParam(0) = "코스트센터팝업"	
			arrParam(1) = "B_COST_CENTER"				
			arrParam(2) = Trim(frm1.txtCostCd.Value)
			arrParam(3) = ""
			arrParam(4) = "cost_cd in (select distinct cost_cd from c_dir_mfc_cost where yyyymm = " & FilterVar(strYYYYMM1, "''", "S") & ")"			
			arrParam(5) = "코스트센터"			
	
			arrField(0) = "cost_cd"	
			arrField(1) = "cost_nm"	
    
			arrHeader(0) = "코스트센터"		
			arrHeader(1) = "코스트센터명"		
		case 2
			arrParam(0) = "계정팝업"	
			arrParam(1) = "A_ACCT"				
			arrParam(2) = Trim(frm1.txtAcctCd.Value)
			arrParam(3) = ""
			arrParam(4) = "acct_cd in (select distinct acct_cd from c_dir_mfc_cost where yyyymm = " & FilterVar(strYYYYMM1, "''", "S") & ")"						
			arrParam(5) = "계정"			
	
			arrField(0) = "acct_cd"	
			arrField(1) = "acct_nm"	
    
			arrHeader(0) = "계정"		
			arrHeader(1) = "계정명"				
		case 3
			arrParam(0) = "관리항목"	
			arrParam(1) = "A_CTRL_ITEM"				
			arrParam(2) = Trim(frm1.txtCtrlCd.Value)
			arrParam(3) = ""
			arrParam(4) = "ctrl_cd in (select distinct ctrl_cd from c_dir_mfc_cost where yyyymm = " & FilterVar(strYYYYMM1, "''", "S") & ")"						
			arrParam(5) = "관리항목"			
	
			arrField(0) = "ctrl_cd"	
			arrField(1) = "ctrl_nm"	
    
			arrHeader(0) = "관리항목"		
			arrHeader(1) = "관리항목명"				
	End select 
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	  select case iWhere
		case 1
			frm1.txtCostCd.focus
		case 2
			frm1.txtAcctCd.focus
		case 3
			frm1.txtCtrlCd.focus
	  End Select	
		Exit Function
	Else
		Call SetPopup(iWhere,arrRet)
	End If	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPopup(byval iWhere,byval arrRet)
	select case iWhere
		case 1
			frm1.txtCostCd.Value    = arrRet(0)		
			frm1.txtCostNm.Value    = arrRet(1)
			frm1.txtCostCd.focus()
		case 2
			frm1.txtAcctCd.Value    = arrRet(0)		
			frm1.txtAcctNm.Value    = arrRet(1)
			frm1.txtAcctCd.focus()		
		case 3
			frm1.txtCtrlCd.Value    = arrRet(0)		
			frm1.txtCtrlNm.Value    = arrRet(1)
			frm1.txtCtrlCd.focus()		
	End select		
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

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet("*")                                               '⊙: Setup the Spread sheet
   
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
 
    frm1.txtYyyymm.focus
	Set gActiveElement = document.activeElement
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    IF frm1.hSpid.value <> "" Then
		lgCloseFlgMode = "C"
		call DBQuery()
	END IF	
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

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


'=======================================================================================================
'   Event Name : txtYyyymm_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub



'==========================================================================================
'   Event Name : vspdData1_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then							'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub



'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    lgOldRow = frm1.vspdData1.ActiveRow
			
	frm1.vspdData2.MaxRows = 0 
					
	If DbDtlQuery(Row) = False Then	
		Call RestoreToolBar()
		Exit Sub
	End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
    Else
        
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
      gMouseClickStatus = "SP2CR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	gMouseClickStatus = "SPC"	'Split 상태코드    
	
    If Row <> NewRow And NewRow > 0 Then
		lgOldRow = frm1.vspdData1.ActiveRow
				
		frm1.vspdData2.MaxRows = 0 
						
		If DbDtlQuery(NewRow) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If	    
	    
	End If    
	    

End Sub


'

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

    FncQuery = False														'⊙: Processing is NG
    Err.Clear																'☜: Protect system from crashing


    IF frm1.hSpid.value <> "" Then
		lgCloseFlgMode = "C"
	
		If DbQuery = False Then Exit Function		
    END IF
    
    
	If frm1.txtCostCd.value = "" Then
		frm1.txtCostNm.value = "" 
	End If	
	
	If frm1.txtAcctCd.value = "" Then
		frm1.txtAcctNm.value = "" 
	End If
	
	
	If frm1.txtCtrlCd.value = "" Then
		frm1.txtCtrlNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables														'⊙: Initializes local global variables
	
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------



    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db data
	End If
	
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
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
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
    On Error Resume Next													'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
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
	FncExit = True
End Function

'******************  5.2 Fnc함수명에서 호출되는 개발 Function  **************************
'	설명 : 
'**************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
    Dim strYear,strMonth,strDay
    
    DbQuery = False

	Call LayerShowHide(1)
    
    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMM = strYear & strMonth

	With frm1
		If lgIntFlgMode <> parent.OPMD_UMODE and lgCloseFlgMode <> "C" Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtYyyymm="			& strYYYYMM			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtCostCd="			& UCase(Trim(.txtCostCd.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtAcctCd="			& UCase(Trim(.txtAcctCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtCtrlCd="			& UCase(Trim(.txtCtrlCd.value))
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
			strVal = strVal & "&txtSpid="			& Trim(.hspid.value)
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0003						'☜: 
			strVal = strVal & "&txtYyyymm="			& Trim(.hYYYYMM.value)			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtCostCd="			& UCase(Trim(.hCostCd.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtAcctCd="			& UCase(Trim(.hAcctCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtCtrlCd="			& UCase(Trim(.hCtrlCd.value))
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
			strVal = strVal & "&txtSpid="			& Trim(.hspid.value)
			

		End If
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")														'⊙: 버튼 툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	
	If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
		Call RestoreToolBar()
		Exit Function
	End If
	
	frm1.hYYYYMM.value = strYYYYMM
		
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal NewRow) 
    
    Dim strVal
       
    DbDtlQuery = False

	Call LayerShowHide(1)
    
	With frm1
		strVal = BIZ_PGM_QRY2_ID & "?txtMode="	& parent.UID_M0001						'☜: 

		.vspdData1.Row = NewRow
		.vspdData1.Col = C_CostCd
		strVal = strVal & "&txtCostCd="			& UCase(Trim(.vspdData1.text))			'☆: 조회 조건 데이타 
	
		.vspdData1.Col = C_AcctCd	
		strVal = strVal & "&txtAcctCd="			& UCase(Trim(.vspdData1.text))			'☆: 조회 조건 데이타 
		
		.vspdData1.Col = C_CtrlCd	
		strVal = strVal & "&txtCtrlCd="			& UCase(Trim(.vspdData1.text))			'☆: 조회 조건 데이타 

		.vspdData1.Col = C_CtrlVal	
		strVal = strVal & "&txtCtrlVal="			& UCase(Trim(.vspdData1.text))			'☆: 조회 조건 데이타 
		
		strVal = strVal & "&txtSpid="			& UCase(Trim(.hSpid.value))		

		strVal = strVal & "&txtCondCostCd="			& UCase(Trim(.hCostCd.value))			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCondAcctCd="			& UCase(Trim(.hAcctCd.value))			'☆: 조회 조건 데이타		
		strVal = strVal & "&txtCondCtrlCd="			& UCase(Trim(.hCtrlCd.value))		
		
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
    
    DbDtlQuery = True
    
End Function


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
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>직과가공비배부내역조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>작업년월</TD> 
									<TD CLASS="TD6">
										<script language =javascript src='./js/c3940ma1_OBJECT1_txtYyyymm.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>코스트센터</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="코스트센터"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=20 tag="14"></TD>

								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계정</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(2)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>관리항목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCtrlCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="관리항목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(3)">&nbsp;<INPUT TYPE=TEXT NAME="txtCtrlNm" SIZE=20 tag="14"></TD>

								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=40% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/c3940ma1_A_vspdData1.js'></script>
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
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>총배부대상금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3940ma1_fpDoubleSingle1_txtSum1.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>  
				<TR>
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/c3940ma1_B_vspdData2.js'></script>
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
								<TD CLASS=TD5>총배부금액</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3940ma1_fpDoubleSingle2_txtSum2.js'></script>&nbsp;
    							<TD CLASS=TD5>배부금액합계</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3940ma1_fpDoubleSingle2_txtSum3.js'></script>&nbsp;
                               </TD>
    							<TD CLASS=TD5>배부기준Data합계</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3940ma1_fpDoubleSingle2_txtSum4.js'></script>&nbsp;
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hSpid" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hAcctCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hCtrlCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

