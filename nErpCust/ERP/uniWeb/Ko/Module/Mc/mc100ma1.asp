<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC100MA1
'*  4. Program Name         : 납입지시대상선정 
'*  5. Program Desc         : 납입지시대상선정 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-04-08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ahn Jung Je
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
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
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =====================================-->

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
Const BIZ_PGM_ID = "mc100mb1.asp"											
Const BIZ_PGM_JUMP_ID = "M3111MA1"
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

Dim C_Select		
Dim C_ProdOrderNo	'제조오더번호 
Dim C_ItemCd		'품목 
Dim C_ItemNm		'품목명 
Dim C_Spec			'규격 
Dim C_ReqDt			'필요일 
Dim C_Unit			'재고단위 
Dim C_ReqQty		'필요수량 
Dim C_PoBaseQty		'발주수량 
Dim C_PoNeedQty		'발주필요수량 
Dim C_PoUnit		'발주단위 
Dim C_PoUnitQty		'발주단위발주수량 
Dim C_TrackingNo	'Tracking No
Dim C_WCCd			'작업장  
Dim C_PlanStartDt	'착수예정일 
Dim C_PlanComptDt	'완료예정일 
Dim C_OprNo			'공정 
Dim C_Seq			'부품예약순서 
Dim C_SubSeq		'납입지시순번 
Dim C_ReleaseDt

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop    
Dim strDate
Dim iDBSYSDate

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtFromReqDt.text	= UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtToReqDt.text	= UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	Call SetToolbar("1110000000001111")
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","M","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030226", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_ReleaseDt + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck C_Select, "선택",8,,,true
		ggoSpread.SSSetEdit 	C_ProdOrderNo,  "제조오더번호"	,20
		ggoSpread.SSSetEdit 	C_ItemCd,       "품목"			,20
		ggoSpread.SSSetEdit 	C_ItemNm,       "품목명"		,25
		ggoSpread.SSSetEdit 	C_Spec,			"규격"			,25
		ggoSpread.SSSetEdit 	C_ReqDt,		"필요일"		,10, 2
		ggoSpread.SSSetEdit 	C_Unit,			"재고단위"		,10
		ggoSpread.SSSetFloat	C_ReqQty,		"필요량"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_PoBaseQty,	"발주수량",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_PoNeedQty,	"발주필요수량"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_PoUnit,		"발주단위"		,10
		ggoSpread.SSSetFloat	C_PoUnitQty,	"발주단위발주수량"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No"	,25
		ggoSpread.SSSetEdit 	C_WcCd,			"작업장"		,20
		ggoSpread.SSSetEdit 	C_PlanStartDt,  "착수예정일"	,11, 2
		ggoSpread.SSSetEdit 	C_PlanComptDt,	"완료예정일"	,11, 2
		ggoSpread.SSSetEdit 	C_OprNo,		"공정"			,12
		ggoSpread.SSSetEdit 	C_Seq,			"부품예약순서"	,12
		ggoSpread.SSSetEdit 	C_SubSeq,		"납입지시순번"	,12
		ggoSpread.SSSetEdit 	C_ReleaseDt,	"작업지시일"	,11, 2
		
		Call ggoSpread.SSSetColHidden(C_Seq, .MaxCols,	True)
		Call SetSpreadLock 
		
		.ReDraw = true    
    End With
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.SpreadLock	-1, -1
    ggoSpread.SpreadUnLock	C_Select, -1, C_Select, -1
End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2110", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDoTime, lgF0, lgF1, Chr(11))
    
    frm1.cboDoTime.value = ""
End Sub

'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()
	C_Select		= 1
	C_ProdOrderNo	= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_ReqDt			= 6
	C_Unit			= 7
	C_ReqQty		= 8
	C_PoBaseQty		= 9
	C_PoNeedQty		= 10
	C_PoUnit		= 11
	C_PoUnitQty		= 12
	C_TrackingNo	= 13
	C_WCCd			= 14
	C_PlanStartDt	= 15
	C_PlanComptDt	= 16
	C_OprNo			= 17
	C_Seq			= 18
	C_SubSeq		= 19
	C_ReleaseDt		= 20
End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_Select		= iCurColumnPos(1)
			C_ProdOrderNo	= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_ReqDt			= iCurColumnPos(6)
			C_Unit			= iCurColumnPos(7)
			C_ReqQty		= iCurColumnPos(8)
			C_PoBaseQty		= iCurColumnPos(9)
			C_PoNeedQty		= iCurColumnPos(10)
			C_PoUnit		= iCurColumnPos(11)
			C_PoUnitQty		= iCurColumnPos(12)
			C_TrackingNo	= iCurColumnPos(13)
			C_WCCd			= iCurColumnPos(14)
			C_PlanStartDt	= iCurColumnPos(15)
			C_PlanComptDt	= iCurColumnPos(16)
			C_OprNo			= iCurColumnPos(17)
			C_Seq			= iCurColumnPos(18)
			C_SubSeq		= iCurColumnPos(19)
			C_ReleaseDt		= iCurColumnPos(20)
    End Select
End Sub    

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
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
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo(i)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If i = 1 then		
		If IsOpenPop = True Or UCase(frm1.txtProdOrderNo1.className) = "PROTECTED" Then Exit Function
	Else
		If IsOpenPop = True Or UCase(frm1.txtProdOrderNo2.className) = "PROTECTED" Then Exit Function
	End if
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	Else
		If Plant_Check() = False Then Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""									'ProdFromDt
	arrParam(2) = ""									'ProdToDt
	arrParam(3) = "RL"									'From Status
	arrParam(4) = "ST"									'To Status
	If i = 1 then	
		arrParam(5) = Trim(frm1.txtProdOrderNo1.value)
	Else
		arrParam(5) = Trim(frm1.txtProdOrderNo2.value)
	End if
	
	arrParam(6) = ""		
	arrParam(7) = ""		
	arrParam(8) = ""			
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If i = 1 then	
			frm1.txtProdOrderNo1.focus
		Else
			frm1.txtProdOrderNo2.focus
		End if	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		If i = 1 then	
			frm1.txtProdOrderNo1.Value	= arrRet(0)
			frm1.txtProdOrderNo1.focus
		Else
			frm1.txtProdOrderNo2.Value	= arrRet(0)
			frm1.txtProdOrderNo2.focus
		End if	
	End If
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	Else
		If Plant_Check() = False Then Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	frm1.vspdData.Row =	frm1.vspdData.ActiveRow 
	

		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'===========================================================================
' Function Name : JumpChgCheck
' Function Desc : Jump시 데이타 변경여부 체크 
'=========================================================================== 
Function JumpChgCheck()
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFromReqDt.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_Select, i, 1)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_Select, i, 0)
		Next	
	End If
End Sub
'=======================================================================================================
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtFromReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtToReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromReqDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToReqDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then Exit Sub							'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
        
    If OldLeft <> NewLeft Then Exit Sub
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then									'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
%>
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0

			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row 
			frm1.vspdData.text = "" 
			lgBlnFlgChgValue = False					
	    End Select
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then Exit Sub                                                    'If there is no data.

   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
    Else
        
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False														'⊙: Processing is NG
    Err.Clear																'☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If

    If Not chkfield(Document, "1") Then	Exit Function									'⊙: This function check indispensable field
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables		
    
	If Trim(frm1.txtItemCd.value) <> "" Then
		If Plant_Item_Check() = False Then Exit Function
	Else
		frm1.txtItemNm.value = ""
		If Plant_Check() = False Then Exit Function
	End If
	
	If ValidDateCheck(frm1.txtFromReqDt, frm1.txtToReqDt) = False Then Exit Function
												'⊙: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'☜: Query db data
	End If
	
	Set gActiveElement = document.activeElement
    FncQuery = True															'⊙: Processing is OK
End Function

'===========================================  5.1.2 FncNew()  ===========================================
'=	Event Name : FncNew																					=
'=	Event Desc : This function is related to New Button of Main ToolBar									=
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False								

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then Exit Function
	End If

	Call ggoOper.ClearField(Document, "A")			
	Call ggoOper.LockField(Document, "N")			
	Call SetDefaultVal
	Call SetToolBar("11100000000011")				

	Call InitVariables
		
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If								

	Set gActiveElement = document.activeElement
	FncNew = True									
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear         
    
    If CheckRunningBizProcess = True Then Exit Function
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then                 
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")   
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                        
    If Not ggoSpread.SSDefaultCheck Then Exit Function         
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                         
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'☜: Protect system from crashing
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		
		If IntRetCD = vbNo Then Exit Function
		
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    Dim strYear1
    Dim strMonth1
    Dim strDay1
    Dim strDate1
       
    DbQuery = False
    
	Call LayerShowHide(1)
    
	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtFromReqDt="		& Trim(.hFromReqDt.value)			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToReqDt="		& Trim(.hToReqDt.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtProdOrderNo1="	& UCase(Trim(.hProdOrderNo1.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtProdOrderNo2="	& UCase(Trim(.hProdOrderNo2.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey							'
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'☆: 조회 조건 데이타 
			strVal = strVal & "&txtFromReqDt="		& Trim(.txtFromReqDt.text)			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtToReqDt="		& Trim(.txtToReqDt.text)			'☆: 조회 조건 데이타 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'☆: 조회 조건 데이타		
			strVal = strVal & "&txtProdOrderNo1="	& UCase(Trim(.txtProdOrderNo1.value))	'☆: 조회 조건 데이타 
			strVal = strVal & "&txtProdOrderNo2="	& UCase(Trim(.txtProdOrderNo2.value))	'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
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
	Call SetToolBar("11101000000111")														'⊙: 버튼 툴바 제어 
	lgIntFlgMode = parent.OPMD_UMODE														'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data Save 
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal
    Dim igColSep,igRowSep
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
    Dim ii

    DbSave = False    

    igColSep = Parent.gColSep
    igRowSep = Parent.gRowSep
	frm1.txtMode.value = Parent.UID_M0002
    lGrpCnt = 1
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
   
    If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		For lRow = 1 To .vspdData.MaxRows
		
			If Trim(GetSpreadText(.vspdData,0,lRow,"X","X")) = ggoSpread.UpdateFlag then
				strVal = "U" & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProdOrderNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X"))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X"))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow,"X","X"))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_ReqDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_ReqQty,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_WCCd,lRow,"X","X"))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PlanStartDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PlanComptDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_ReleaseDt,lRow,"X","X")))    & igColSep
				strVal = strVal & lRow & igRowSep

				lGrpCnt = lGrpCnt + 1
			End If
			
			Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
			    Case ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
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
			End Select   
		Next
	End With
	
	frm1.txtMaxRows.value = lGrpCnt - 1
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                       
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()										
	Call InitVariables()
	Call MainQuery()
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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
    ggoSpread.Source = gActiveSpdSheet
		
	Call ggoSpread.ReOrderingSpreadData()
End Sub 

'========================================================================================
' Function Name : Plant_Check
' Function Desc : 
'========================================================================================
Function Plant_Check()
	Plant_Check = False

	'-----------------------
	'Check Plant CODE		'공장코드가 있는 지 체크 
	'-----------------------
    If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus 
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	Plant_Check = True
End Function

'========================================================================================
' Function Name : Plant_Item_Check
' Function Desc : 
'========================================================================================
Function Plant_Item_Check()
	Plant_Item_Check = False

	'-----------------------
	'Check Item CODE		'공장코드가 있는 지 체크 
	'-----------------------
    If 	CommonQueryRs(" C.PLANT_NM, B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B, B_PLANT C ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD " & _
						" AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
			Call DisplayMsgBox("122700","X","X","X")
			frm1.txtItemCd.focus 
		Else
			Call DisplayMsgBox("122600","X","X","X")
			frm1.txtItemNm.Value = ""
			frm1.txtItemCd.focus 
		End If
		
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	frm1.txtItemNm.Value = lgF1(0)
	
	Plant_Item_Check = True
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시대상 선정</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" ALT="공장명"></TD>
									<TD CLASS=TD5 NOWRAP>필요일</TD> 
									<TD CLASS="TD6">
										<script language =javascript src='./js/mc100ma1_OBJECT1_txtFromReqDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/mc100ma1_OBJECT2_txtToReqDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td NOWRAP>
													<INPUT TYPE=TEXT NAME="txtProdOrderNo1" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo(1)"></TD>
												<td NOWRAP> &nbsp;~ &nbsp;</td>
											</tr>
										</table>
									</TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MaxLength=40 tag="14"></TD>
									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo2" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo(2)"></TD>
									
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
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <script language =javascript src='./js/mc100ma1_OBJECT3_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
	    <td WIDTH="100%">
	    	<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="left">
					    <button name="btnSelect" class="clsmbtn">일괄선택</button>&nbsp;
					    <BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</td>
					<TD WIDTH="*" ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">발주등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo1" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo2" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

