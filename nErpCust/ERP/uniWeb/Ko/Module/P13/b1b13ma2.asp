<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Component Allocation Entry
'*  3. Program ID           : b1b13ma2.asp
'*  4. Program Name         : 대체품조회 
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/03/14
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Mr  Kim GyoungDon
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "b1b13mb3.asp"								'☆: 비지니스 로직 ASP명 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "b1b13mb4.asp"								'☆: 비지니스 로직 ASP명 

' Grid 1(vspdData1) - Operation 
Dim C_ITEM_CD			
Dim C_ITEM_NM			
Dim C_SPECIFICATION
Dim C_ITEM_ACCT			
Dim C_ITEM_CLASS		
Dim C_PROCUR_TYPE		
Dim C_BASIC_UNIT		
Dim C_PRODT_ENV			
Dim C_PHANTOM_FLG		
Dim C_MPS_FLAG			
Dim C_TRACKING_FLG		
Dim C_ORDER_TYPE		
Dim C_ORDER_RULE		
Dim C_LOT_FLG			
Dim C_VALID_FLG			
Dim C_VALID_FROM_DT1	
Dim C_VALID_TO_DT1		
                     
' Grid 2(vspdData2) - Operation 
Dim C_ALT_ITEM_CD		
Dim C_ALT_ITEM_NM		
Dim C_ALT_ITEM_SPEC
Dim C_PRIORITY		
Dim C_VALID_FROM_DT2
Dim C_VALID_TO_DT2	
Dim C_SEQ			

Dim BaseDate
Dim StartDate
Dim strYear
Dim strMonth
Dim strDay

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)
Call ExtractDateFrom(BaseDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim lgStrPrevKey2

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop 
Dim lgLngCnt
Dim lgOldRow
         
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey2 = ""
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow = 0
    lgSortKey = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDt.Text = StartDate
	frm1.txtToDt.Text = UniConvYYYYMMDDToDate(Parent.gDateFormat, "2999","12","31")
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub


'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdId)
	
	Call InitSpreadPosVariables(pvSpdId)

	If pvSpdId = "*" Or pvSpdId = "A" Then
	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
		With frm1.vspdData1 
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021122", , Parent.gAllowDragDropSpread

			.ReDraw = False
   
			.MaxCols = C_VALID_TO_DT1 + 1
			.MaxRows = 0    

			Call GetSpreadColumnPos("A")
	    
			ggoSpread.SSSetEdit C_ITEM_CD,		"품목",			15					
			ggoSpread.SSSetEdit C_ITEM_NM,		"품목명",		25					
			ggoSpread.SSSetEdit C_SPECIFICATION,"규격",			25
			ggoSpread.SSSetEdit C_ITEM_ACCT,	"품목계정",		12					
			ggoSpread.SSSetEdit C_ITEM_CLASS,	"집계용품목클래스",14				
			ggoSpread.SSSetEdit C_PROCUR_TYPE,	"조달구분",		10
			ggoSpread.SSSetEdit C_BASIC_UNIT,	"단위",			8					
			ggoSpread.SSSetEdit C_PRODT_ENV,	"생산전략",		10
			ggoSpread.SSSetEdit C_PHANTOM_FLG,	"팬텀",			10, 2				
			ggoSpread.SSSetEdit C_MPS_FLAG,		"MPS구분",		10, 2				
			ggoSpread.SSSetEdit C_TRACKING_FLG,	"Tracking여부",	10, 2
			ggoSpread.SSSetEdit C_ORDER_TYPE,	"오더생성여부",	10, 2											    
			ggoSpread.SSSetEdit C_ORDER_RULE,	"Lot Sizing",	12
			ggoSpread.SSSetEdit C_LOT_FLG,		"LOT관리",		10, 2			
			ggoSpread.SSSetEdit C_VALID_FLG,	"유효구분",		10, 2		
			ggoSpread.SSSetDate	C_VALID_FROM_DT1,"시작일",		12, 2, Parent.gDateFormat
			ggoSpread.SSSetDate	C_VALID_TO_DT1, "종료일",		12, 2, Parent.gDateFormat

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

			ggoSpread.SSSetSplit(1)											'frozen 기능 추가 
   
			.ReDraw = True
    
		End With
	End If
	
	If pvSpdId = "*" Or pvSpdId = "B" Then
	'------------------------------------------
	' Grid 2 - Component Spread Setting
	'------------------------------------------
		With frm1.vspdData2

    		ggoSpread.Source = frm1.vspdData2
     		ggoSpread.Spreadinit "V20021122", , Parent.gAllowDragDropSpread    

			.ReDraw = False

			.MaxCols = C_SEQ + 1
			.MaxRows = 0
    
			Call GetSpreadColumnPos("B")
  
		    ggoSpread.SSSetEdit	C_ALT_ITEM_CD,		"대체품목",		15
		    ggoSpread.SSSetEdit	C_ALT_ITEM_NM,		"대체품목명",	25
		    ggoSpread.SSSetEdit	C_ALT_ITEM_SPEC,	"대체품목규격",	25  
		    ggoSpread.SSSetEdit	C_PRIORITY ,		"우선순위",		10 ,1
		    ggoSpread.SSSetDate C_VALID_FROM_DT2,	"시작일",		12, 2, Parent.gDateFormat
		    ggoSpread.SSSetDate C_VALID_TO_DT2,		"종료일",		12, 2, Parent.gDateFormat
			ggoSpread.SSSetEdit	C_SEQ,				"순서",			6,2

			Call ggoSpread.SSSetColHidden(C_SEQ, C_SEQ, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit(1)											'frozen 기능 추가 
	
			.ReDraw = True
    
		End With
	End If
	
	Call SetSpreadLock(pvSpdId)
	    
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables(ByVal pvSpdId)

	If pvSpdId = "*" Or pvSpdId = "A" Then
	' Grid 1(vspdData1) - Operation 
		C_ITEM_CD			= 1  
		C_ITEM_NM			= 2  
		C_SPECIFICATION		= 3 
		C_ITEM_ACCT			= 4  
		C_ITEM_CLASS		= 5  
		C_PROCUR_TYPE		= 6  
		C_BASIC_UNIT		= 7  
		C_PRODT_ENV			= 8  
		C_PHANTOM_FLG		= 9  
		C_MPS_FLAG			= 10     
		C_TRACKING_FLG		= 11 
		C_ORDER_TYPE		= 12 
		C_ORDER_RULE		= 13 
		C_LOT_FLG			= 14 
		C_VALID_FLG			= 15 
		C_VALID_FROM_DT1	= 16 
		C_VALID_TO_DT1		= 17 
	End If

	If pvSpdId = "*" Or pvSpdId = "B" Then
	' Grid 2(vspdData2) - Operation                          
		C_ALT_ITEM_CD		= 1     
		C_ALT_ITEM_NM		= 2     
		C_ALT_ITEM_SPEC		= 3     
		C_PRIORITY			= 4        
		C_VALID_FROM_DT2	= 5     
		C_VALID_TO_DT2		= 6     
		C_SEQ				= 7
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Dim IntRetCD
	
	Set gActiveSpdSheet = frm1.vspdData1
	gMouseClickStatus = "SPC"
	Call SetPopupMenuItemInf("0000111111") 
   
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
		
		Exit Sub
		
	End If
	
	If  Col < 0 Then
		Exit Sub
	End If
	
	'----------------------
	'Column Split
	'----------------------

	If lgOldRow <> Row Then
		
		lgOldRow = Row
		frm1.vspdData2.MaxRows = 0
		LayerShowHide(1)
		
		Call DisableToolBar(Parent.TBC_QUERY)   ': Query 버튼을 disable 시킴.
        
        If DbDtlQuery = False Then 
           Call RestoreToolBar()
           Exit Sub
        End If 			

	End If
	
End Sub


'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

	'----------------------
	'Column Split
	'----------------------
	Set gActiveSpdSheet = frm1.vspdData2
	gMouseClickStatus = "SP2C"   
	Call SetPopupMenuItemInf("0000111111") 
   
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

'========================================================================================
' Function Name : vspdData1_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData1
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
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet(gActiveSpdSheet.id)
   Call ggoSpread.ReOrderingSpreadData
End Sub 


'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData1
				
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				
		C_ITEM_CD			= iCurColumnPos(1)      
		C_ITEM_NM			= iCurColumnPos(2)      
		C_SPECIFICATION		= iCurColumnPos(3)     
		C_ITEM_ACCT			= iCurColumnPos(4)      
		C_ITEM_CLASS		= iCurColumnPos(5)      
		C_PROCUR_TYPE		= iCurColumnPos(6)      
		C_BASIC_UNIT		= iCurColumnPos(7)      
		C_PRODT_ENV			= iCurColumnPos(8)      
		C_PHANTOM_FLG		= iCurColumnPos(9)      
		C_MPS_FLAG			= iCurColumnPos(10)      
		C_TRACKING_FLG		= iCurColumnPos(11)     
		C_ORDER_TYPE		= iCurColumnPos(12)     
		C_ORDER_RULE		= iCurColumnPos(13)     
		C_LOT_FLG			= iCurColumnPos(14)     
		C_VALID_FLG			= iCurColumnPos(15)     
		C_VALID_FROM_DT1	= iCurColumnPos(16)     
		C_VALID_TO_DT1		= iCurColumnPos(17)     
		
	Case "B"
		ggoSpread.Source = frm1.vspdData2 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_ALT_ITEM_CD		= iCurColumnPos(1) 
		C_ALT_ITEM_NM		= iCurColumnPos(2) 
		C_ALT_ITEM_SPEC		= iCurColumnPos(3) 
		C_PRIORITY			= iCurColumnPos(4) 
		C_VALID_FROM_DT2	= iCurColumnPos(5) 
		C_VALID_TO_DT2		= iCurColumnPos(6) 
		C_SEQ				= iCurColumnPos(7) 
	End Select
End Sub


'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock(ByVal pvSpdId)

    With frm1
		If pvSpdId = "*" Or pvSpdId = "A" Then	
			'--------------------------------
			'Grid 1
			'--------------------------------
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
    
		If pvSpdId = "*" Or pvSpdId = "B" Then	
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
	
    End With
End Sub


'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal lRow)

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()
	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Item Account
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))

	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Procurement Type(조달구분)
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
End Sub

'------------------------------------------  OpenConItemInfo()  -------------------------------------------------
'	Name : OpenConItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenConPlant()  -----------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

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

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                                    '⊙: Setup the Spread sheet
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    Call InitVariables                                                      '⊙: Initializes local global variables
    
    Call SetToolbar("11000000000011")
    
    If Parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
		
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement  
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
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
'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData1_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData1_onfocus()

End Sub

'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData2_onfocus()

End Sub


'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
   
End Sub 

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

   If Button = 2 And gMouseClickStatus = "SP2C" Then
      gMouseClickStatus = "SP2CR"
   End If
End Sub 



'=======================================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	
End Sub

'=======================================================================================================
'   Event Name : vspdData2_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	
	If lgOldRow <> Row Then
		lgLngCurRows = NewRow
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData2_Change(ByVal Col , ByVal Row )
	
End Sub


'==========================================================================================
'   Event Name : vspdData_DragDropBlock
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData2_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================


Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

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
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
		If lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			LayerShowHide(1)
			Call DisableToolBar(Parent.TBC_QUERY)   ': Query 버튼을 disable 시킴.
            If DbDtlQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 	
		End If     
    End if
    
End Sub


'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
 

Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)

End Sub


Sub txtPlantCd_OnChange()
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
End Sub

Sub txtItemCd_OnChange()
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
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
    
    FncQuery = False															'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing
   
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
   
   	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function       
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     														'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
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
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 

End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow()
 
	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 

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
    Call parent.FncExport(Parent.C_SINGLEMULTI)                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                                    '☜: Protect system from crashing
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
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing
        
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtItemAcct=" & Trim(.hItemAcct.value)
		strVal = strVal & "&txtProcType=" & Trim(.hProcType.value)
		strVal = strVal & "&txtFromDt=" & Trim(.hFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(.hToDt.value)
		strVal = strVal & "&rdoValidFlg=" & Trim(.hValidFlg.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		
    Else
	
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtItemAcct=" & Trim(.cboAccount.value)
		strVal = strVal & "&txtProcType=" & Trim(.cboProcType.value)
		strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
		If .rdoValidFlg1.checked = True Then
			strVal = strVal & "&rdoValidFlg=" & Trim(.rdoValidFlg1.value)
		ElseIf .rdoValidFlg2.checked = True Then
			strVal = strVal & "&rdoValidFlg=" & Trim(.rdoValidFlg2.value)
		Else 
			strVal = strVal & "&rdoValidFlg=" & Trim(.rdoValidFlg3.value)
		End If
		
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		
    End If
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================


Function DbQueryOk()
	
	Call SetToolbar("11000000000111")

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisableToolBar(Parent.TBC_QUERY)   ': Query 버튼을 disable 시킴.
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
		End If
        If DbDtlQuery = False Then 
           Call RestoreToolBar()
           Exit Function
        End If 	
	End If

	lgIntFlgMode = Parent.OPMD_UMODE
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbDtlQuery() 
    Dim strVal
    Dim SelItemCd
    
    DbDtlQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing
	
	'frm1.vspdData1.Col = 1
	frm1.vspdData1.Col = C_ITEM_CD
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow 
	
	SelItemCd = Trim(frm1.vspdData1.Text)
		        
    With frm1
    
	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001						'☜: 
	strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtItemCd=" & Trim(SelItemCd)
	strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False

End Function


'========================================================================================
' Function Name : DbQuery
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>대체품조회</font></td>
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
			 						<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목계정</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="계정" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcType" ALT="조달구분" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>종료일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/b1b13ma2_I912441056_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/b1b13ma2_I112995517_txtToDt.js'></script>					
									</TD>
									<TD CLASS=TD5 NOWRAP>유효구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" tag="11" CHECKED ID="rdoValidFlg1" VALUE="A"><LABEL FOR="rdoValidFlg1">전체</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" tag="11" ID="rdoValidFlg2" VALUE="Y"><LABEL FOR="rdoValidFlg2">예</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidFlg" tag="11" ID="rdoValidFlg3" VALUE="N"><LABEL FOR="rdoValidFlg3">아니오</LABEL></TD>
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
							<TR HEIGHT="100%">
								<TD WIDTH="35%">
									<script language =javascript src='./js/b1b13ma2_A_vspdData1.js'></script>
								</TD>
								<TD WIDTH="65%">
									<script language =javascript src='./js/b1b13ma2_B_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24"><INPUT TYPE=HIDDEN NAME="hProcType" tag="24"><INPUT TYPE=HIDDEN NAME="hFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hValidFlg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
