<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b25ma1.asp
'*  4. Program Name         : Copy Item by Plant
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2004/03/19
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Park In Sik
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance
On Error Resume Next
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY_ID	= "B3B25MB1.asp"												'☆: Detail Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "B3B25MB2.asp"												'☆: Detail Query 비지니스 로직 ASP명 

'==========================================================================================================
'==========================================================================================================

Dim C_Select 
Dim C_Item
Dim C_ItmNm
Dim C_ItmSpec
Dim C_PrcCtrlInd
Dim C_PrcCtrlIndNm
Dim C_UnitPrice
Dim C_IBPValidFromDt
Dim C_IBPValidToDt
Dim C_ClassCd       
Dim C_ClassDesc     
Dim C_CharValue1    
Dim C_CharValueDesc1
Dim C_CharValue2    
Dim C_CharValueDesc2
Dim C_ItmAcc
Dim C_HdnItmAcc
Dim C_Unit
Dim C_Phantom
Dim C_ItmGroupCd
Dim C_ItmGroupNm
Dim C_DefaultFlg
Dim C_ValidFromDt
Dim C_ValidToDt

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgInsrtFlg
Dim lgFlgAllSelected		'When Selected All
Dim lgFlgCancelClicked		'Cancel Button Clicked
Dim lgFlgCopyClicked		'Copy Button Clicked
Dim lgFlgBtnSelectAllClicked 'When btnSelectAll Clicked

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop					 'Popup
Dim iDBSYSDate
Dim StartDate, EndDate

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_Select          		= 1
	C_Item            		= 2
	C_ItmNm           		= 3
	C_ItmSpec         		= 4
	C_PrcCtrlInd      		= 5
	C_PrcCtrlIndNm    		= 6
	C_UnitPrice       		= 7
	C_IBPValidFromDt  		= 8	
	C_IBPValidToDt    		= 9
	C_ClassCd         		= 10
	C_ClassDesc       		= 11
	C_CharValue1      		= 12
	C_CharValueDesc1  		= 13
	C_CharValue2      		= 14
	C_CharValueDesc2  		= 15 
	C_ItmAcc          		= 16
	C_HdnItmAcc       		= 17
	C_Unit            		= 18
	C_Phantom				= 19 
	C_ItmGroupCd      		= 20
	C_ItmGroupNm      		= 21
	C_DefaultFlg      		= 22
	C_ValidFromDt     		= 23
	C_ValidToDt       		= 24 

End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	frm1.btnCopy.disabled = True
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "전체선택"
	lgFlgAllSelected = False
	lgFlgCancelClicked = False
	lgFlgCopyClicked = False
	lgFlgBtnSelectAllClicked = False

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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
    With frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030208",,parent.gAllowDragDropSpread   
	
    .MaxCols = C_ValidToDt + 1											'☜: 최대 Columns +1
    .MaxRows = 0
    
	.ReDraw = false

 	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetCheck	C_Select ,		"",					2,,,1
	ggoSpread.SSSetEdit 	C_Item,			"품목",			20,,,18,2
	ggoSpread.SSSetEdit 	C_ItmNm,		"품목명",		25,,,40
	ggoSpread.SSSetEdit 	C_ItmSpec,		"규격",			25,,,40
	ggoSpread.SSSetCombo 	C_PrcCtrlInd,	"단가구분",		12
	ggoSpread.SSSetCombo 	C_PrcCtrlIndNm, "단가구분",		20
	ggoSpread.SSSetFloat	C_UnitPrice,	"단가",			15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_IBPValidFromDt,"유효시작일",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_IBPValidToDt,	"유효종료일",	12, 2, parent.gDateFormat
	ggoSpread.SSSetEdit 	C_ClassCd,		"클래스",		20,,,18,2
	ggoSpread.SSSetEdit 	C_ClassDesc,	"클래스명",		25,,,40
	ggoSpread.SSSetEdit 	C_CharValue1,	"사양값1",		20,,,18,2
	ggoSpread.SSSetEdit 	C_CharValueDesc1,"사양값명1",	25,,,40
	ggoSpread.SSSetEdit 	C_CharValue2,	"사양값2",		20,,,18,2
	ggoSpread.SSSetEdit 	C_CharValueDesc2,"사양값명2",	25,,,40
	ggoSpread.SSSetCombo 	C_ItmAcc,		"품목계정",		12
	ggoSpread.SSSetCombo 	C_HdnItmAcc,	"품목계정",		16
	ggoSpread.SSSetEdit 	C_Unit,			"기준단위",		10,,,3,2
	ggoSpread.SSSetEdit 	C_Phantom,		"팬텀",			10,2
	ggoSpread.SSSetEdit 	C_ItmGroupCd,	"품목그룹",		18,,,10,2
	ggoSpread.SSSetEdit 	C_ItmGroupNm,	"품목그룹명",	25
	ggoSpread.SSSetEdit 	C_DefaultFlg,	"유효구분",		8,2
	ggoSpread.SSSetDate		C_ValidFromDt,	"품목시작일",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ValidToDt,	"품목종료일",	12, 2, parent.gDateFormat
	
	Call ggoSpread.SSSetColHidden(C_HdnItmAcc,	C_HdnItmAcc,	True)
	Call ggoSpread.SSSetColHidden(C_PrcCtrlInd,	C_PrcCtrlInd,	True)
	
	Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,		True)
	
	ggoSpread.SSSetSplit2(2)										'frozen 기능추가 
	Call SetSpreadLock 
	
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

 	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLock -1, -1
	ggoSpread.spreadUnLock C_Select, -1, C_Select

End Sub

'================================== 2.2.5 SetSpreadLock1() ==================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles When An Specific Row is Selected
'=============================================================================================================
Sub SetSpreadLock1(ByVal Col, ByVal Row)

 	ggoSpread.SpreadLock		C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadLock		C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadLock		C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
	
	ggoSpread.SSSetProtected	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetProtected	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetProtected	C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetProtected	C_IBPValidToDt,		Row, Row

End Sub

'================================== 2.2.6 SetSpreadUnLock() ==================================================
' Function Name : SetSpreadUnLock
' Function Desc : This method set color and protect in spread sheet celles When A Specific Row is Selected
'=============================================================================================================
Sub SetSpreadUnLock(ByVal Col, ByVal Row)

	ggoSpread.SpreadUnLock		C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadUnLock		C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadUnLock		C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadUnLock		C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
    
	ggoSpread.SSSetRequired 	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetRequired		C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetRequired		C_IBPValidToDt,		Row, Row

End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Select         		= iCurColumnPos(1)
			C_Item           		= iCurColumnPos(2)
			C_ItmNm          		= iCurColumnPos(3)
			C_ItmSpec        		= iCurColumnPos(4)
			C_PrcCtrlInd     		= iCurColumnPos(5)
			C_PrcCtrlIndNm   		= iCurColumnPos(6)
			C_UnitPrice      		= iCurColumnPos(7)
			C_IBPValidFromDt 		= iCurColumnPos(8)	
			C_IBPValidToDt   		= iCurColumnPos(9)
			C_ClassCd        		= iCurColumnPos(10)
			C_ClassDesc      		= iCurColumnPos(11)
			C_CharValue1     		= iCurColumnPos(12)
			C_CharValueDesc1 		= iCurColumnPos(13)
			C_CharValue2     		= iCurColumnPos(14)
			C_CharValueDesc2 		= iCurColumnPos(15)
			C_ItmAcc         		= iCurColumnPos(16)
			C_HdnItmAcc      		= iCurColumnPos(17)
			C_Unit           		= iCurColumnPos(18)
			C_Phantom	     		= iCurColumnPos(19)
			C_ItmGroupCd     		= iCurColumnPos(20)
			C_ItmGroupNm     		= iCurColumnPos(21)
			C_DefaultFlg     		= iCurColumnPos(22)
			C_ValidFromDt    		= iCurColumnPos(23)
			C_ValidToDt      		= iCurColumnPos(24)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm
	
	'****************************
    ' 품목계정 
    '****************************     
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc
    	
	'****************************
    ' 팬텀,통합구매,유효구분 
    '****************************    
    strCboCd = ""
    strCboNm = ""
    
    strCboCd = "Y" & vbTab & "N"
		
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting 
    	
	'****************************
    'Price Control Ind
    '****************************
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & "S" & vbTab				'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "표준단가" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & "M"						'& vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "이동평균단가"		'& vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
    Dim strCboCd 
    Dim strCboNm
    
    '****************************
    ' 품목계정 
    '****************************    
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc
    
	'****************************
    ' 팬텀,통합구매,유효구분 
    '****************************    
    strCboCd = ""
    strCboNm = ""
    
    strCboCd = "Y" & vbTab & "N"
		
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting 
	 	
	'****************************
    'Price Control Ind
    '****************************	
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & UCase("S") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "표준단가" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & UCase("M") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "이동평균단가" & vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		'.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			If iPos = 1 Then
				.Row = intRow
				.Col = C_HdnItmAcc
				intIndex = .value
				.col = C_ItmAcc
				.value = intindex

				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
				
			Else
				.Row = intRow
				.Col = C_ItmAcc
				intIndex = .value
				.col = C_HdnItmAcc
				.value = intindex
			
				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
			End IF							
		Next	
		'.ReDraw = True
	End With
End Sub

Function SetFieldProp(ByVal lRow, ByVal sType)
	ggoSpread.Source = frm1.vspdData
    
	ggoSpread.SSSetRequired	C_PrcCtrlInd,	lRow, lRow
	ggoSpread.SSSetRequired	C_UnitPrice,	lRow, lRow
End Function

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

'------------------------------------------  OpenCondPlant()  ---------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
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
		Call SetConPlant(arrRet, 0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(strCode, iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If iPos = 1 Then
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("971012","X", "공장","x")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	End If
	
	IsOpenPop = True
	
	If iPos = 0 Then
		arrParam(0) = strCode						' Item Code
		arrParam(1) = ""							' Item Name
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B01PA2")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
			IsOpenPop = False
			Exit Function
		End If
	ElseIf iPos = 1 Then
		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
		arrParam(1) = strCode						' Item Code
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B11PA4")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
			IsOpenPop = False
			Exit Function
		End If
	End If

    arrField(0) = 1 								' Field명(0) :"ITEM_CD"
    arrField(1) = 2 								' Field명(1) :"ITEM_NM"
    arrField(2) = 3 								' Field명(2) :"SPEC"
    arrField(3) = 9 								' Field명(2) :"ProcurType"
    arrField(4) = 10 								' Field명(2) :"ProcurType"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet,iPos)
	End If	

	Call SetFocusToDocument("M")
	If iPos = "0" Then
		frm1.txtItemCd.focus
	Else
		frm1.txtItemCd1.focus
	End If	

End Function

'------------------------------------------  OpenItemGroup()  ---------------------------------------------
'	Name : OpenItemGroup()
'	Description : Condition Item Group PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = frm1.txtHighItemGroupCd.value  
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "품목그룹"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
'    arrField(3) = "LEAF_FLG"	
'    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroupCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtHighItemGroupCd.focus
	
End Function

'------------------------------------------  OpenClassCd()  -------------------------------------------------
'	Name : OpenClassCd()
'	Description : Class PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenClassCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtClasscd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtClassCd.value)	' Class Code
	arrParam(1) = ""							' Class Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "Class_CD"
    arrField(1) = 2 							' Field명(1) : "Class_NM"
	
	iCalledAspName = AskPRAspName("B3B31PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B31PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetClassCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtClassCd.focus
	
End Function

'==========================================================================================
'   Event Name : txtItemCd1_onChange()
'   Event Desc :
'==========================================================================================
Sub txtItemCd1_onChange()
	With frm1
		If .txtItemCd1.value = "" Then
			.txtItemNm1.value = ""
			.txtItemSpec1.value = ""
			.txtItemProcType1.value = ""
	
			.txtItemCd1.focus
			Set gActiveElement = document.activeElement
		Else	
			Call LookUpItemByPlant()
		End If
	End With
End Sub

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlant()
	Dim iStrWhereSQL
	Dim strITEM_CD
	Dim strITEM_NM
	Dim strSPEC
	Dim strPROCUR_TYPE_CD
	Dim strPROCUR_TYPE_NM

	iStrWhereSQL = "A.ITEM_CD = B.ITEM_CD AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd1.value, "''", "S") & " AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	Call CommonQueryRs(" A.ITEM_CD, A.ITEM_NM, A.SPEC, B.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & " , B.PROCUR_TYPE) "," B_ITEM A, B_ITEM_BY_PLANT B ",iStrWhereSQL ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	strITEM_CD = lgF0
	strITEM_NM = lgF1
	strSPEC = lgF2
	strPROCUR_TYPE_CD = lgF3
	strPROCUR_TYPE_NM = lgF4
		
	strITEM_CD			=	replace(strITEM_CD,Chr(11),"")
	strITEM_NM			=	replace(strITEM_NM,Chr(11),"")
	strSPEC				=	replace(strSPEC,Chr(11),"")
	strPROCUR_TYPE_CD	=	replace(strPROCUR_TYPE_CD,Chr(11),"")
	strPROCUR_TYPE_NM	=	replace(strPROCUR_TYPE_NM,Chr(11),"")
	
	frm1.txtItemNm1.value = strITEM_NM
	frm1.txtItemSpec1.value = strSPEC
	frm1.txtItemProcType1.value = strPROCUR_TYPE_NM
	frm1.htxtItemProcType1.value = strPROCUR_TYPE_CD
End Function

'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	frm1.txtClassCd.Value    = arrRet(0)		
	frm1.txtClassNm.Value    = arrRet(1)
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet,ByVal iPos)
	With frm1
		If iPos = 0 Then
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)
		Else
			.txtItemCd1.value	= arrRet(0)
			.txtItemNm1.value	= arrRet(1)
			.txtItemSpec1.value = arrRet(2)
			.txtItemProcType1.value = arrRet(4)
			.htxtItemProcType1.value = arrRet(3)
		End If

	End With
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet, byval iPos)
	With frm1
		If iPos = 0 Then
			.txtPlantCd.Value    = arrRet(0)		
			.txtPlantNm.Value    = arrRet(1)
		Else
			.txtPlantCd1.Value    = arrRet(0)		
		End If
	End With
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_Unit
	frm1.vspdData.Text = arrRet(0)
End Function

'------------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItmGroupCd
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_ItmGroupNm
	frm1.vspdData.Text = arrRet(1)		
End Function

'------------------------------------------  SetItemGroupCd()  --------------------------------------------------
'	Name : SetItemGroupCd()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroupCd(Byval arrRet)
	frm1.txtHighItemGroupCd.value = arrRet(0)
	frm1.txtHighItemGroupNm.value = arrRet(1)
End Function

'------------------------------------------  SetBaseItem()  --------------------------------------------------
'	Name : SetBaseItem()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBasisItemCd(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_BaseItm
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_BaseItmNm
	frm1.vspdData.Text = arrRet(1)		

End Function

'------------------------------------------  ChkBtnAll()  --------------------------------------------------
'	Name : ChkBtnAll()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function btnSelectAll_Clicked()
	Dim LngRow
	
	If frm1.vspdData.MaxRows <= 0 Then Exit Function

	lgFlgBtnSelectAllClicked = True
	frm1.btnSelectAll.disabled = True
	
	With frm1.vspdData
		
		.ReDraw = False

		If lgFlgAllSelected = False Then 'select all clicked
				
			For LngRow = 1 To .MaxRows
				Call .SetText(C_Select,LngRow,"1")
				Call SetSpreadUnLock(C_Select, LngRow)
				If lgInsrtFlg <> True Then
					ggoSpread.UpdateRow LngRow
				End If
			Next

			Call InitData(1,1)	
			
			frm1.btnSelectAll.value = "전체선택취소"
			lgFlgAllSelected = True
			
		Else 'deselect all clicked
			
			For LngRow = 1 To .MaxRows
				If GetSpreadText(frm1.vspdData,C_Select,LngRow,"X","X") = "1" _
				And GetSpreadText(frm1.vspdData,0,LngRow,"X","X") <> ggoSpread.InsertFlag Then
					Call .SetText(C_Select,LngRow,"0")
					Call ggoSpread.EditUndo(LngRow, LngRow)
					Call SetSpreadLock1(C_Select, LngRow)
				End If
			Next
			
			Call InitData(1,1)

			frm1.btnSelectAll.value = "전체선택"
			lgFlgAllSelected = False
		End If
		.ReDraw = True
	End With
	
	frm1.btnSelectAll.disabled = False
	lgFlgBtnSelectAllClicked = False		

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
	Err.Clear
	
	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet															'⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal	'함수 정의가 없음 
	Call InitVariables	'함수 정의가 없음											'⊙: Initializes local global variables
	
	Call SetToolbar("11000000000011")												'⊙: 버튼 툴바 제어 
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtPlantCd1.value = parent.gPlant
			frm1.txtItemCd.focus
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
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("0000110111")
		Else
			Call SetPopupMenuItemInf("0001111111") 
		End if			
	End If	
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData  
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row )
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.Col = C_Select
		.Row = Row

		If .Value = "1" Then
			ggoSpread.UpdateRow Row
		End if
		
		If Col = C_PrcCtrlIndNm Then
		   Call vspdData_ComboSelChange (C_PrcCtrlIndNm,Row)
		End If   

	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp 
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------   

	If frm1.vspdData.Row <= 0 Or lgFlgBtnSelectAllClicked = True Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		If gMouseClickStatus = "SPC" Or lgFlgCancelClicked = True Then
			If Col = C_Select And Not (lgFlgCopyClicked) Then
				If GetSpreadText(frm1.vspdData,C_Select,Row,"X","X") = "0" Then
					.Redraw = false
					Call SetSpreadLock1(C_Select, Row)
					call ggoSpread.EditUndo(Row, Row)
					Call InitData(1,1)
					.Redraw = true
				Else
					.Redraw = false
					Call SetSpreadUnLock(C_Select, Row)	
					.Redraw = true
				End If
			End If
		End If
				
		Select Case Col
			Case C_Select
				If lgInsrtFlg <> True Then
					If Buttondown = 1 Then
						ggoSpread.Source = frm1.vspdData
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.Source = frm1.vspdData
						ggoSpread.SSDeleteFlag Row,Row
					End If
				End If
		End Select
    End With
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  C_ItmAcc
				.Col = Col
				intIndex = .Value
				.Col = C_HdnItmAcc
				.Value = intIndex
			'Case  C_SumItmClass
			'	.Col = Col
			'	intIndex = .Value
			'	.Col = C_HdnSumItmClass
			'	.Value = intIndex
			Case  C_PrcCtrlIndNm
				.Col = Col
				intIndex = .value
				.Col = C_PrcCtrlInd
				.value = intIndex
				If Trim(.Text) = "M" Then
					ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetProtected 	C_UnitPrice, 		Row, Row
				Else
					ggoSpread.SpreadUnLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
				End If
						
		End Select
    
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then Exit Sub
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    'If CheckRunningBizProcess = True Then
	'	Exit Sub
	'End If	
	
	'----------  Coding part  ------------------------------------------------------------- 				
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크	
    	If lgStrPrevKey <> "" Then		                                            '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)									': Query 버튼을 disable 시킴.
			frm1.vspdData.ReDraw = False
			If DbQuery = False Then
				Call RestoreToolBar()
				frm1.vspdData.ReDraw = True
				Exit Sub
			End if
			frm1.vspdData.ReDraw = True
    	End If
    End If

End Sub
'==========================================================================================
'   Event Name : txtPlantCd_OnChange
'   Event Desc : This function is Setting the txtPlantCd,txtPlantNm
'==========================================================================================
Sub txtPlantCd_OnBlur()
	With frm1
		If Trim(.txtPlantCd.value) = "" Then
			.txtPlantNm.value = ""
			.txtPlantCd1.value = ""
		Else
			.txtPlantCd1.value = UCase(Trim(.txtPlantCd.value))
		End If
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
    Dim strPlantCd
    Dim strPlantNm
    Dim strPlantItem
    Dim strPlantItemNm
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

	'-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    	
	Call ggoOper.ClearField(Document, "3")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
	
	 '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	
	If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	Else
		strPlantItem = frm1.txtItemCd1.value 
		strPlantItemNm = frm1.txtItemNm1.value 
	End If
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	    
    'Call SetDefaultVal
	
	If strPlantItem <> "" Then    
		frm1.txtItemCd1.value = strPlantItem
		frm1.txtItemNm1.value = strPlantItemNm
	End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

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
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	End If
	    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Precheck area
    '-----------------------
   	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear
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

	If frm1.vspdData.MaxRows <= 0 Then Exit Function
	
	lgFlgCancelClicked = True
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.ReDraw = False
		Call ggoSpread.EditUndo(.ActiveRow,.ActiveRow)
		Call InitData(1,1)
		Call SetSpreadLock1(C_Select, .ActiveRow)
		.ReDraw = True
	End With
	
	lgFlgCancelClicked = False
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
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()	
	Call InitData(1,1)
End Sub


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
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
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '⊙: Processing is NG
    
    Call LayerShowHide(1)
    
    Dim strVal
	
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001								'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)					'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtClassCd=" & Trim(frm1.htxtClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)				'☆: 조회 조건 데이타 
		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtHighItemGroupCd.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)				'☆: 조회 조건 데이타 
			
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    End If
		
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
    
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
    Call InitData(LngMaxRow,1)
	
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field
    
    Call SetToolbar("11000001000111")
    
    frm1.btnCopy.disabled = True
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "전체선택"
	lgFlgAllSelected = False    
    
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim IntRows 
    Dim lGrpcnt 
    Dim strVal
    Dim GenVal
	Dim IntRetCD
	Dim iColSep
	Dim TmpBuffer
	Dim iTotalStr
	
	DbSave = False														'⊙: Processing is NG
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002									'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    ReDim TmpBuffer(0)
    iColSep = Parent.gColSep 
	GenVal = "10000000"
	
	If UCase(frm1.htxtItemProcType1.value) = "M" Then
		
	End If	
	
	With frm1.vspdData

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
	
			Select Case .Text
				Case ggoSpread.UpdateFlag
					
					strVal = ""
						
					strVal = strVal & "C" & iColSep	& IntRows & iColSep				'⊙: U=Update
					
					.Col = C_Item								'2
					strVal = strVal & Trim(.Text) & iColSep

					.Col = C_PrcCtrlInd							'3
					strVal = strVal & Trim(.Text) & iColSep		
					
					If (Trim(UCase(frm1.htxtItemProcType1.value)) = "M" And Trim(UCase(.Text)) = "M") Or (Trim(UCase(frm1.htxtItemProcType1.value)) = "O" And Trim(UCase(.Text)) = "M") Then
						IntRetCD = DisplayMsgBox("122726", parent.VB_INFORMATION, "X", "X")	'조달구분이 사내가공품이면 단가구분은 표준단가만 가능합니다.
						Call LayerShowHide(0)
						Exit Function
					End If
					
					.Col = C_UnitPrice							'4
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep		

					.Col = C_Phantom							'5
			        strVal = strVal & Trim(.Text) & iColSep

					.Col = C_IBPValidFromDt						'6
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep						
						    
					.Col = C_IBPValidToDt						'7
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

					.Col = C_ValidFromDt						'8
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep						
						    
					.Col = C_ValidToDt						'9
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

					strVal = strVal & GenVal & parent.gRowSep			'10			'⊙: 마지막 데이타는 Row 분리기호를 넣는다		        
					
					ReDim Preserve TmpBuffer(lGrpcnt-1)
					
					TmpBuffer(lGrpcnt-1) = strVal
							
					lGrpcnt = lGrpcnt + 1             
			End Select
	   Next
	
	End With
	
	iTotalStr = Join(TmpBuffer, "")
	
	frm1.txtMaxRows.value = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = iTotalStr										'☜: Spread Sheet 내용을 저장 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 

    DbSave = True                                                           '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables

    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0

    Call FncQuery()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<!--'#########################################################################################################
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공장별품목정보 COPY</font></td>
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
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="24"></TD>
									<TD CLASS=TD5 NOWRAP>클래스</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtClassCd" SIZE=20 MAXLENGTH=16 tag="11XXXU"  ALT="클래스"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()">
										<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value,0">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14">
									</TD>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=20 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btHighItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >
										<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=20 tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>기준품목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="22XXXU" ALT="기준품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd1.value,1"></TD>
								<TD CLASS="TD5" NOWRAP>기준품목규격</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec1" SIZE=40 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>기준품목명</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>기준품목조달구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemProcType1" SIZE=40 tag="24"></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="32" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD>
						<BUTTON NAME="btnCopy" CLASS="CLSMBTN" Flag=1 ONCLICK="FncSave">COPY</BUTTON>&nbsp;
						<BUTTON NAME="btnSelectAll" CLASS="CLSMBTN" Flag=1 ONCLICK="btnSelectAll_Clicked">전체선택</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TabIndex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hItemAccount" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hPhantomFlg" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hItemClass" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="htxtClassCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtPlantCd1" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="htxtItemProcType1" tag="24" TabIndex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>
