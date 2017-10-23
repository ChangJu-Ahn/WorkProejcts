<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : Preliminary Delivery Order Status
'*  3. Program ID           : mc900qb1
'*  4. Program Name         : 납입지시대상조회 
'*  5. Program Desc         : List Preliminary Delivery Order Status
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/03/05
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Seung Wook
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "mc900qb1.asp"								'☆: Head Query 비지니스 로직 ASP명 

Dim C_PlantCd
Dim C_PlantNm
Dim C_PoNo 		
Dim C_PoSeqNo 
Dim C_ItemCd					
Dim C_ItemNm					
Dim C_ItemSpec	
Dim C_DlvyDt	
Dim C_PoQty
Dim C_PoUnit
Dim C_BaseQty		
Dim C_BaseUnit	
Dim C_PoDlyQty		 
Dim C_PoRcptQty
Dim C_BaseDlyQty
Dim C_BaseRcptQty
Dim C_SlCd		 
Dim C_BpCd		 
Dim C_BpNm
Dim C_PurOrg
Dim C_PurGrp
Dim C_PoDt
Dim C_PrNo
Dim C_TrackingNo
Dim C_ProcureType

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop										'Popup
Dim lgStrPrevKey1,lgStrPrevKey2

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False					'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""	
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboDlvyOrderFlag, "C", "생성")
	Call SetCombo(frm1.cboDlvyOrderFlag, "I", "진행")
	Call SetCombo(frm1.cboDlvyOrderFlag, "F", "완료")
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
	Dim LocSvrDate
	
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtPoFrDt.text	  = UniConvDateAToB(UNIDateAdd ("D", -5, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtPoToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	Call SetToolbar("1100000000001111")	
End Sub
   
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()
    
    With frm1
    
	    ggoSpread.Source = .vspdData
	    ggoSpread.Spreadinit "V20030305", , Parent.gAllowDragDropSpread

	 	.vspdData.ReDraw = false
	    .vspdData.MaxCols = C_ProcureType + 1
	    .vspdData.MaxRows = 0

	    Call GetSpreadColumnPos("A")

	    ggoSpread.SSSetEdit		C_PlantCd,		"공장", 10
	    ggoSpread.SSSetEdit		C_PlantNm,		"공장명", 20
	    ggoSpread.SSSetEdit 	C_PoNo,			"발주번호", 20
		ggoSpread.SSSetEdit 	C_PoSeqNo,		"발주순번", 10,1
	    ggoSpread.SSSetEdit		C_ItemCd,		"품목", 20,,,,2
	    ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 25
	    ggoSpread.SSSetEdit		C_ItemSpec,		"규격", 25
	    ggoSpread.SSSetDate 	C_DlvyDt,		"납기일", 12, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat	C_PoQty,		"발주량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_PoUnit,		"발주단위", 10
		ggoSpread.SSSetFloat	C_BaseQty,		"재고단위발주수량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_BaseUnit,		"재고단위", 10
	    ggoSpread.SSSetFloat	C_PoDlyQty,		"발주단위 납입지시가능량",22,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_PoRcptQty,	"발주단위 입고량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_BaseDlyQty,	"재고단위 납입지시가능량",22,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_BaseRcptQty,	"재고단위 입고량",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_SlCd,			"창고", 10
	    ggoSpread.SSSetEdit		C_BpCd,			"공급처", 10
	    ggoSpread.SSSetEdit		C_BpNm,			"공급처명", 20
	    ggoSpread.SSSetEdit		C_PurOrg,		"구매조직", 8
	    ggoSpread.SSSetEdit		C_PurGrp,		"구매그룹", 8
	    ggoSpread.SSSetDate 	C_PoDt,			"발주일", 12, 2, parent.gDateFormat
	    ggoSpread.SSSetEdit		C_PrNo,			"구매요청번호", 20
	    ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
	    ggoSpread.SSSetEdit		C_ProcureType,	"조달구분", 8
	        
   
	    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		
		.vspdData.ReDraw = true
		
	    ggoSpread.Source = .vspdData
    End With

    Call SetSpreadLock()
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
  ggoSpread.Source = frm1.vspdData
  ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData 
    
		.Redraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetProtected C_PlantCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PlantNm, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoNo ,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoSeqNo , 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCd, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DlvyDt, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoQty,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoUnit, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BaseQty, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BaseUnit,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoDlyQty, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoRcptQty,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BaseDlyQty,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BaseRcptQty,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SlCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BpCd, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BpNm, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PurOrg,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PurGrp, 			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoDt,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PrNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProcureType,		pvStartRow, pvEndRow

		.Redraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd						= 1
	C_PlantNm						= 2
	C_PoNo							= 3
	C_PoSeqNo						= 4
	C_ItemCd						= 5	
	C_ItemNm						= 6
	C_ItemSpec						= 7
	C_DlvyDt						= 8
	C_PoQty							= 9 
	C_PoUnit						= 10
	C_BaseQty						= 11
	C_BaseUnit						= 12
	C_PoDlyQty						= 13
	C_PoRcptQty						= 14
	C_BaseDlyQty					= 15
	C_BaseRcptQty					= 16
	C_SlCd							= 17	
	C_BpCd							= 18
	C_BpNm							= 19
	C_PurOrg						= 20
	C_PurGrp						= 21
	C_PoDt							= 22
	C_PrNo							= 23
	C_TrackingNo					= 24
	C_ProcureType					= 25		
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
			
			C_PlantCd					= iCurColumnPos(1)
			C_PlantNm					= iCurColumnPos(2)
			C_PoNo						= iCurColumnPos(3)
			C_PoSeqNo					= iCurColumnPos(4)  
			C_ItemCd					= iCurColumnPos(5)  
			C_ItemNm					= iCurColumnPos(6)  
			C_ItemSpec					= iCurColumnPos(7)  
			C_DlvyDt					= iCurColumnPos(8)  
			C_PoQty						= iCurColumnPos(9)  
			C_PoUnit					= iCurColumnPos(10) 
			C_BaseQty					= iCurColumnPos(11) 
			C_BaseUnit					= iCurColumnPos(12) 
			C_PoDlyQty					= iCurColumnPos(13) 
			C_PoRcptQty					= iCurColumnPos(14) 
			C_BaseDlyQty				= iCurColumnPos(15) 
			C_BaseRcptQty				= iCurColumnPos(16) 
			C_SlCd						= iCurColumnPos(17) 
			C_BpCd						= iCurColumnPos(18) 
			C_BpNm						= iCurColumnPos(19) 
			C_PurOrg					= iCurColumnPos(20) 
			C_PurGrp					= iCurColumnPos(21) 
			C_PoDt						= iCurColumnPos(22) 
			C_PrNo						= iCurColumnPos(23) 
			C_TrackingNo				= iCurColumnPos(24) 
			C_ProcureType				= iCurColumnPos(25) 
			
 	End Select
End Sub
 
'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
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

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim iCalledAspName
	Dim arrParam(5), arrField(2)
	
	Dim IntRetCD
	Dim arrRet

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    '공장 체크 함수 호출 
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","X")
		IsOpenPop = False
		Exit Function
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

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
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

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Biz Partner PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtBpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus    	
		Exit Function
	Else
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)	
		frm1.txtBpCd.focus    	
		Set gActiveElement = document.activeElement	
	End If	
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field
    Call InitSpreadSheet 

	Call InitVariables		'⊙: Initializes local global variables

	 'Plant Code, Plant Name Setting 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	
		frm1.txtBpCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
	
	Call SetDefaultVal
	Call InitComboBox
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
         
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Sub
		If Row < 1 Then
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
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
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
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtPoFrDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPoFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPoToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPoToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPoFrDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtPoFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtPoToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtPoToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = "" 
	End If
	
    If ValidDateCheck(frm1.txtPoFrDt, frm1.txtPoToDt) = False Then Exit Function
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkfield(Document, "1") Then											'⊙: This function check indispensable field
'       Exit Function
'    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function										'☜: Query db data
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                           '☜: Protect system from crashing
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'☜: Protect system from crashing
    Set gActiveElement = document.ActiveElement   
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
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Err.Clear							'☜: Protect system from crashing

    DbQuery = False                                                         			'⊙: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
   
    If frm1.txtPlantCd.value <> "" Then
		If CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			LayerShowHide(0)
			Exit Function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	Else 
		frm1.txtPlantNm.value = ""
	End if
	
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				Call DisplayMsgBox("122600","X","X","X")
				frm1.txtItemNm.Value = ""
				LayerShowHide(0)
				Exit Function
			Else
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtItemNm.Value = lgF0(0)
				Call DisplayMsgBox("122700","X","X","X")
				LayerShowHide(0)
				Exit Function
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtItemNm.Value = lgF0(0)
	Else
		frm1.txtItemNm.Value = ""
	End if
	
	If frm1.txtBpCd.value <> "" Then
		If CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", " BP_CD = " & FilterVar(frm1.txtBpCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("179021","X","X","X")
			frm1.txtBpNm.Value = ""
			LayerShowHide(0)
			Exit Function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtBpNm.Value = lgF0(0)
	Else
		frm1.txtBpNm.Value = ""
	End if
 
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode="				& Parent.UID_M0001						
		strVal = strVal			& "&lgIntFlgMode="			& lgIntFlgMode		
		strVal = strVal			& "&lgStrPrevKey1="			& lgStrPrevKey1  
		strVal = strVal			& "&lgStrPrevKey2="			& lgStrPrevKey2  
		strVal = strVal			& "&txtPlantCd="			& Trim(frm1.hPlantCd.value)			
		strVal = strVal			& "&txtPoFrDt="				& Trim(frm1.hPoFrDt.value)		
		strVal = strVal			& "&txtPoToDt="				& Trim(frm1.hPoToDt.value)			
		strVal = strVal			& "&txtItemCd="				& Trim(frm1.hItemCd.value)				
		strVal = strVal			& "&txtBpCd="				& Trim(frm1.hBpCd.value)
		strVal = strVal			& "&cboDlvyOrderFlag="		& Trim(frm1.hDlvyOrderFlag.value)
		
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode="				& Parent.UID_M0001					
		strVal = strVal			& "&lgIntFlgMode="			& lgIntFlgMode
		strVal = strVal			& "&lgStrPrevKey1="			& lgStrPrevKey1
		strVal = strVal			& "&lgStrPrevKey2="			& lgStrPrevKey2
		strVal = strVal			& "&txtPlantCd="			& Trim(frm1.txtPlantCd.value)			
		strVal = strVal			& "&txtPoFrDt="				& Trim(frm1.txtPoFrDt.text)		
		strVal = strVal			& "&txtPoToDt="				& Trim(frm1.txtPoToDt.text)			
		strVal = strVal			& "&txtItemCd="				& Trim(frm1.txtItemCd.value)				
		strVal = strVal			& "&txtBpCd="				& Trim(frm1.txtBpCd.value)
		strVal = strVal			& "&cboDlvyOrderFlag="		& Trim(frm1.cboDlvyOrderFlag.value)							
	End If    

    Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'☆: 조회 성공후 실행로직 
	Call SetToolBar("11000000000111")											'⊙: 버튼 툴바 제어 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field

	If frm1.vspdData.MaxRows <= 0 Then Exit Function
    lgIntFlgMode = Parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시대상조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="공급처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>					   
								</TR>					   
								<TR>						   
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
														   <INPUT TYPE=TEXT Alt="품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE cellspacing=0 cellpadding=0>
											<TD>
												<TD>
													<script language =javascript src='./js/mc900qa1_fpDateTime2_txtPoFrDt.js'></script>
												</TD>
												<TD>&nbsp;~&nbsp;</TD>
												<TD>
													<script language =javascript src='./js/mc900qa1_fpDateTime2_txtPoToDt.js'></script>
												</TD>
											<TD>
										</TABLE>
							         </TD>
	                            </TR>
	                            <TR>
									<TD CLASS="TD5" NOWRAP>진행여부</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDlvyOrderFlag" ALT="진행여부" STYLE="Width: 120px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
	                            </TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/mc900qa1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD<%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hPoFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hPoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hDlvyOrderFlag" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hPoSeqNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
