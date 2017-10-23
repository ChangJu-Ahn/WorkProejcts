<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name			: 공정별원가 
'*  2. Function Name		: 
'*  3. Program ID			: c4007ma1.asp
'*  4. Program Name			: 원부자재그룹별원가요소등록 
'*  5. Program Desc			:
'*  6. Business ASP List	: 
'*  7. Modified date(First)	: 2005/09/12
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: 
'* 10. Modifier (Last)		: HJO
'* 11. Comment		: 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID = "c4007mb1.asp"			'☆: Head Query 비지니스 로직 ASP명 

Dim C_GroupLevel 
Dim C_GroupLevelPopup
Dim C_ItemGroup
Dim C_ItemGroupPopup 
Dim C_ItemGroupNM

Dim C_CostElmtCd
Dim C_CostElmtPopup 
Dim C_CostElmtNM

Dim C_ComCostElmtCd
Dim C_ComCostElmtPopup 
Dim C_ComCostElmtNM
	

Dim BaseDate
Dim StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
	lgSortKey = 1
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
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
     
   	Call InitSpreadPosVariables()

    With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021123", , Parent.gAllowDragDropSpread

		.ReDraw = False
	
		.MaxCols = C_ComCostElmtNM + 1
		.MaxRows = 0
    
		Call AppendNumberPlace("6","3","0")

		Call GetSpreadColumnPos("A")
       
		ggoSpread.SSSetEdit		C_GroupLevel,	"그룹Level", 12,,,5,2
		ggoSpread.SSSetButton 	C_GroupLevelPopup 
		ggoSpread.SSSetEdit		C_ItemGroup,	"품목그룹", 10,,,10  
		ggoSpread.SSSetButton		C_ItemGroupPopup 
		ggoSpread.SSSetEdit		C_ItemGroupNM,		"품목그룹명", 25
		ggoSpread.SSSetEdit 		C_CostElmtCd,	"원가요소1",12,,,10  
		ggoSpread.SSSetButton 	C_CostElmtPopup
		ggoSpread.SSSetEdit		C_CostElmtNM,			"원가요소명1", 20
		ggoSpread.SSSetEdit 		C_ComCostElmtCd,	"원가요소2",12,,,10  
		ggoSpread.SSSetButton 	C_ComCostElmtPopup
		ggoSpread.SSSetEdit		C_ComCostElmtNM,			"원가요소명2", 20
		
		Call ggoSpread.MakePairsColumn(C_GroupLevel, C_GroupLevelPopup)
		Call ggoSpread.MakePairsColumn(C_ItemGroup, C_ItemGroupPopup )
		Call ggoSpread.MakePairsColumn(C_CostElmtCd, C_CostElmtPopup )
		Call ggoSpread.MakePairsColumn(C_ComCostElmtCd, C_ComCostElmtPopup )
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		ggoSpread.SSSetSplit2(4)										'frozen 기능추가 
				
		Call SetSpreadLock 

		.ReDraw = True

    End With
    
End Sub


'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()

	C_GroupLevel		= 1
	C_GroupLevelPopup	= 2
	C_ItemGroup			= 3
	C_ItemGroupPopup	= 4
	C_ItemGroupNM		= 5
	C_CostElmtCd		= 6
	C_CostElmtPopup		= 7
	C_CostElmtNM		= 8
	C_ComCostElmtCd		= 9
	C_ComCostElmtPopup	= 10
	C_ComCostElmtNM		= 11
End Sub



'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_GroupLevel		= iCurColumnPos(1)
		C_GroupLevelPopup	= iCurColumnPos(2)
		C_ItemGroup			= iCurColumnPos(3)
		C_ItemGroupPopup	= iCurColumnPos(4)
		C_ItemGroupNM		= iCurColumnPos(5)
		C_CostElmtCd		= iCurColumnPos(6)
		C_CostElmtPopup		= iCurColumnPos(7)
		C_CostElmtNM		= iCurColumnPos(8)
		C_ComCostElmtCd		= iCurColumnPos(9)
		C_ComCostElmtPopup	= iCurColumnPos(10)
		C_ComCostElmtNM		= iCurColumnPos(11)
	End Select

End Sub



'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim IntRetCD

	'Call SetPopupMenuItemInf("1101111111")	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("1001111111")
		Else
			Call SetPopupMenuItemInf("1101111111") 
		End if			
	End If	
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows = 0 Or Col < 0 Then
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
	
	'------ Developer Coding part (Start)
	'------ Developer Coding part (End)
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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
   
End Sub 

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
  
     With frm1

    .vspdData.ReDraw = False
	
	ggoSpread.SSSetRequired		C_ItemGroup,	-1			
	ggoSpread.SpreadLock		C_ItemGroupNM,	-1, C_ItemGroupNM
	ggoSpread.SSSetRequired		C_CostElmtCd,		-1
	ggoSpread.SSSetProtected	C_ComCostElmtNM, -1
	ggoSpread.SSSetProtected	.vspdData.MaxCols, -1
	
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
	
		ggoSpread.SSSetRequired  C_ItemGroup,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemGroupNM,	pvStartRow, pvEndRow
		
		ggoSpread.SSSetRequired  C_CostElmtCd ,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_CostElmtNM, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_ComCostElmtNM, pvStartRow, pvEndRow
		
		.vspdData.ReDraw = True
    
    End With
End Sub
'================================== 2.2.5 SetQuerySpreadColor() ==================================================
' Function Name : SetQuerySpreadColor
' Function Desc :  This method set color and protect  in spread sheet celles, after Query
'========================================================================================

Sub SetQuerySpreadColor()
    
    With frm1
		.vspdData.ReDraw = False
  
		ggoSpread.SSSetProtected C_GroupLevel , -1, -1
		ggoSpread.SSSetProtected C_GroupLevelPopup, -1, -1
		ggoSpread.SSSetProtected C_ItemGroup, -1, -1
		ggoSpread.SSSetProtected C_ItemGroupPopup , -1, -1
		ggoSpread.SSSetProtected C_ItemGroupNM, -1, -1
		ggoSpread.SSSetRequired C_CostElmtCd, -1, -1
		ggoSpread.SSSetProtected C_CostElmtNM, -1, -1
		ggoSpread.SSSetProtected C_ComCostElmtNM, -1, -1		
		
		.vspdData.ReDraw = True
	End With
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


'------------------------------------------  OpenPopup()  -------------------------------------------------
'	Name : OpenPopup()
'	Description : OpenPopup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopup(ByVal strCol, ByVal strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(5)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemGroup.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	Select Case strCol
		Case C_GroupLevelPopup		
			arrParam(0) = "그룹Level"						' 팝업 명칭 
			arrParam(1) = " ( SELECT DISTINCT GROUP_LEVEL FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE 명칭 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "그룹Level"							' TextBox 명칭	
	
			arrField(0) = "ED12" & Parent.gColSep & "group_level"					' Field명(1)
			     
			arrHeader(0) = "그룹Level"						' Header명(0)

		
		Case C_CostElmtPopup 
			arrParam(0) = "원가요소1팝업"						' 팝업 명칭 
			arrParam(1) = " C_COST_ELMT_S  "			' TABLE 명칭 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = " cost_elmt_type='M' "			' Where Condition
			arrParam(5) = "원가요소"							' TextBox 명칭 
	
			arrField(0) = "ED10" & Parent.gColSep & "cost_elmt_cd"					' Field명(1)
			arrField(1) = "ED25" & Parent.gColSep & "cost_elmt_nm"					' Field명(0)
			     
			arrHeader(0) = "원가요소"						' Header명(0)
			arrHeader(1) = "원가요소명"						' Header명(0)   			

		Case C_ComCostElmtPopup 
			arrParam(0) = "원가요소2팝업"						' 팝업 명칭 
			arrParam(1) = " C_COST_ELMT_S  "			' TABLE 명칭 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = " cost_elmt_type='M' "			' Where Condition
			arrParam(5) = "원가요소"							' TextBox 명칭 
	
			arrField(0) = "ED10" & Parent.gColSep & "cost_elmt_cd"					' Field명(1)
			arrField(1) = "ED25" & Parent.gColSep & "cost_elmt_nm"					' Field명(0)
			     
			arrHeader(0) = "원가요소"						' Header명(0)
			arrHeader(1) = "원가요소명"						' Header명(0)   		
		Case C_ItemGroupPopup
			frm1.vspdData.Col = C_GroupLevel  : 			 frm1.vspdData.Row = frm1.vspdData.ActiveRow
			

			arrParam(0) = "품목그룹팝업"						' 팝업 명칭 
			arrParam(1) = " UFN_C_GET_ITEMGROUP() "			' TABLE 명칭 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			If Trim(frm1.vspdData.Text) <>"" Then
			arrParam(4) = " group_level = "			& filterVar(	Trim(frm1.vspdData.Text),"''","S") 	& " "		' Where Condition
			Else
			arrParam(4) = ""							' Where Condition
			End If
			arrParam(5) = "품목그룹"							' TextBox 명칭 
	
	
			arrField(0) = "HH" & Parent.gColSep & "item_group_cd"					' Field명(0)
			arrField(1) = "ED12" & Parent.gColSep & "group_level"					' Field명(1)
			arrField(2) = "ED12" & Parent.gColSep & "item_group_cd"					' Field명(0)
			arrField(3) = "ED20" & Parent.gColSep & "item_group_nm"					' Field명(1)

			     
			arrHeader(0) = "품목그룹"						' Header명(0)
			arrHeader(1) = "그룹Level"						' Header명(1)
			arrHeader(2) = "품목그룹"						' Header명(0)
			arrHeader(3) = "품목그룹명"						' Header명(1)
			
		Case Else
			arrParam(0) = "품목그룹팝업"						' 팝업 명칭 
			arrParam(1) = " UFN_C_GET_ITEMGROUP() "			' TABLE 명칭 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "품목그룹"							' TextBox 명칭 
	
	
			arrField(0) = "HH" & Parent.gColSep & "item_group_cd"					' Field명(0)
			arrField(1) = "ED12" & Parent.gColSep & "group_level"					' Field명(1)
			arrField(2) = "ED12" & Parent.gColSep & "item_group_cd"					' Field명(0)
			arrField(3) = "ED20" & Parent.gColSep & "item_group_nm"					' Field명(1)

			     
			arrHeader(0) = "품목그룹"						' Header명(0)
			arrHeader(1) = "그룹Level"						' Header명(1)
			arrHeader(2) = "품목그룹"						' Header명(0)
			arrHeader(3) = "품목그룹명"						' Header명(1)
					
		
		End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
				
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetValueInfo(strCol, arrRet)
	End If	
End Function

'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetItemPopup()  --------------------------------------------------
'	Name : SetItemPopup()
'	Description : OpenItemPopup Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetValueInfo(Byval strCol, Byval arrRet)
	With frm1
	Select Case strCol
	
	Case C_CostElmtPopup 
			 .vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = C_CostElmtCd
			.vspdData.Text = arrRet(0)	
			.vspdData.Col = C_CostElmtNM
			.vspdData.Text = arrRet(1)			
								
			Call vspdData_Change(strCol, .vspdData.Row)
	Case C_ComCostElmtPopup 
			 .vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = C_ComCostElmtCd
			.vspdData.Text = arrRet(0)	
			.vspdData.Col = C_ComCostElmtNM
			.vspdData.Text = arrRet(1)			
								
			Call vspdData_Change(strCol, .vspdData.Row)			
	Case C_ItemGroupPopup
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = C_GroupLevel
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_ItemGroup
			.vspdData.Text = arrRet(2)			
			.vspdData.Col = C_ItemGroupNm
			.vspdData.Text = arrRet(3)		
			
			Call vspdData_Change(strCol, .vspdData.Row)
	Case C_GroupLevelPopup
			 .vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = C_GroupLevel
			.vspdData.Text = arrRet(0)	
	Case Else

			.txtItemGroup.value = arrRet(2)
			.txtItemGroupNm.value = arrRet(3)		

		Call SetFocusToDocument("M")
		frm1.txtItemGroup.focus
	END SELECT
	End With

End Function



'===========================================================================================================
' Description : checkCode ;check valid code
'===========================================================================================================
Function checkCode(ByVal pvLngRow,byVal pvLngCol ,  ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrCodeInf
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim tmpTxt
	
	checkCode = False
	
	 iStrSelectList="" :  iStrFromList="" : iStrWhereList=""
	With frm1.vspdData
		Select Case pvLngCol
		Case C_GroupLevel
			iStrSelectList = " group_level "
			iStrFromList=" ( SELECT DISTINCT GROUP_LEVEL FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE 명칭 
			iStrWhereList = " group_level =" &  filtervar(pvStrData, "''","S")
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","그룹Level","X")
				frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = pvLngCol : frm1.vspdData.Text =""
				Call SetActiveCell(frm1.vspdData,pvLngCol,pvLngRow,"M","X","X")			
				checkCode = False
				Exit Function
			End If	
'			With frm1.vspdData
				iArrCodeInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = pvLngCol	:  .text = Trim(iArrCodeInf(0))			
'			End With

	
		Case C_ItemGroup
			.Col = C_GroupLevel :tmpTxt = trim(.Text)
			
			iStrSelectList = " item_group_nm  "
			iStrFromList=" ( SELECT DISTINCT GROUP_LEVEL, ITEM_GROUP_CD, ITEM_GROUP_NM,  UPPER_ITEM_GROUP_CD FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE 명칭 
			iStrWhereList ="  item_group_cd =" &  filtervar(pvStrData, "''","S")
			if tmpTxt<>"" then 
			iStrWhereList =iStrWhereList &  " and  group_level =" &  filtervar(tmpTxt, "''","S")			
			End if 
			
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X",frm1.txtItemGroup.alt,"X")
				checkCode = False
				frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_ItemGroupNM : frm1.vspdData.Text =""
				Call SetActiveCell(frm1.vspdData,pvLngCol,pvLngRow,"M","X","X")							
				Exit Function
			End If	
'			With frm1.vspdData
				iArrCodeInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = 	C_ItemGroupNM:  .text = Trim(iArrCodeInf(0))			
'			End With	
		
	
		Case C_CostElmtCd
			iStrSelectList = " cost_elmt_nm   "
			iStrFromList=" C_COST_ELMT_S "			' TABLE 명칭 
			iStrWhereList = " cost_elmt_cd =" &  filtervar(pvStrData, "''","S")
			iStrWhereList = iStrWhereList & " and cost_elmt_type='M'"
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","원가요소1","X")
				frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_CostElmtNM : frm1.vspdData.Text =""
				Call SetActiveCell(frm1.vspdData,pvLngCol,pvLngRow,"M","X","X")			
				checkCode = False
				Exit Function
			End If	
'			With frm1.vspdData
				iArrCodeInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = 	C_CostElmtNM:  .text = Trim(iArrCodeInf(0))			
'			End With	
		Case C_ComCostElmtCd
			iStrSelectList = " cost_elmt_nm   "
			iStrFromList=" C_COST_ELMT_S "			' TABLE 명칭 
			iStrWhereList = " cost_elmt_cd =" &  filtervar(pvStrData, "''","S")
			iStrWhereList = iStrWhereList & " and cost_elmt_type='M'"
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","원가요소2","X")
				frm1.vspdData.Row=pvLngRow :frm1.vspdData.Col = C_ComCostElmtNM : frm1.vspdData.Text =""
				Call SetActiveCell(frm1.vspdData,pvLngCol,pvLngRow,"M","X","X")			
				checkCode = False
				Exit Function
			End If	
'			With frm1.vspdData
				iArrCodeInf = split(lgF0,chr(11))
				.Row = pvLngRow
				.Col = 	C_ComCostElmtNM:  .text = Trim(iArrCodeInf(0))			
'			End With	
		End Select
		
		checkCode = True
		
	End With

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
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field

	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables

	
	'----------  Coding part  -------------------------------------------------------------	
	Call SetToolbar("11001111001111")										'⊙: 버튼 툴바 제어	
   
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
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	with frm1.vspdData
		.Col = Col
		.Row = Row
		Select Case Col
		Case C_GroupLevel
			Call checkCode(Row,Col, .Text)    
		Case C_ItemGroup    
		    Call checkCode(Row, Col, .Text)
		Case C_CostElmtCd    
		    Call checkCode(Row, Col, .Text)
		Case C_ComCostElmtCd    
		    Call checkCode(Row, Col, .Text)
		End Select
	End With
    
End Sub


'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )
     
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()

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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
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

	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
    If Row <1 Then Exit Sub
   
    Select  case Col
		Case  C_ItemGroupPopup 
			.Col = C_ItemGroup
			.Row = Row

			Call OpenPopup (C_ItemGroupPopup,.Text)
        
			Call SetActiveCell(frm1.vspdData,C_CostElmtCd,Row,"M","X","X")
			Set gActiveElement = document.activeElement
		Case C_GroupLevelPopup
			.Col = C_GroupLevelPopup
			.Row = Row

			Call OpenPopup (C_GroupLevelPopup,.Text)
        
			Call SetActiveCell(frm1.vspdData,C_ItemGroup,Row,"M","X","X")
			Set gActiveElement = document.activeElement
		
		Case  C_CostElmtPopup 
			.Col = C_CostElmtPopup
			.Row = Row

			Call OpenPopup (C_CostElmtPopup,.Text)
        
			Call SetActiveCell(frm1.vspdData,C_ComCostElmtCd,Row,"M","X","X")                                                                                                                                                                                                                                                                           
			Set gActiveElement = document.activeElement
		Case  C_ComCostElmtPopup 
			.Col = C_ComCostElmtPopup
			.Row = Row

			Call OpenPopup (C_ComCostElmtPopup,.Text)
        
			Call SetActiveCell(frm1.vspdData,C_ComCostElmtCd,Row,"M","X","X")                                                                                                                                                                                                                                                                           
			Set gActiveElement = document.activeElement
     End Select
    
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
	
    FncQuery = False															'⊙: Processing is NG

    Err.Clear																    '☜: Protect system from crashing
	
	IF ChkKeyField()=False Then Exit Function 
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	If frm1.txtItemGroup.value = "" Then
		frm1.txtItemGroupNm.value = ""
	End If
	  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
    Call InitVariables
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     												'☜: Query db data

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
    Dim IntRetCD 
    Dim iRow
    Dim starDate
    Dim finaDate
    
    FncSave = False																'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing
    On Error Resume Next														'☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    
    If Not ggoSpread.SSDefaultCheck Then              '⊙: Check required field(Multi area)
		Exit Function
    End If  
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
			Exit Function
		End If			
    End If
    
      For iRow=1  to frm1.vspdData.MaxRows			
        frm1.vspdData.Row = iRow
        frm1.vspdData.Col = 0			
		Select Case frm1.vspdData.Text
			Case ggoSpread.InsertFlag	
				frm1.vspdData.Col = C_GroupLevel
				If frm1.vspdData.Text <> "" Then
					If   checkCode(iRow,C_GroupLevel, frm1.vspdData.Text) =False Then Exit Function 					
				End If
							
				frm1.vspdData.Col = C_ITemGroup				
				If  checkCode(iRow,C_ITemGroup, frm1.vspdData.Text) =False Then Exit Function
				
				frm1.vspdData.Col = C_CostElmtCd
				If  checkCode(iRow,C_CostElmtCd, frm1.vspdData.Text) =False Then Exit Function 
				
			'	frm1.vspdData.Col = C_ComCostElmtCd
			'	If  checkCode(iRow,C_ComCostElmtCd, frm1.vspdData.Text) =False Then Exit Function 
							
		End Select	
	Next

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     																				'☜: Save db data
    
    FncSave = True																'⊙: Processing is OK
           
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    frm1.vspdData.focus 
    Set gActiveElement = document.activeElement    
	'frm1.vspdData.EditMode = True
	    
	frm1.vspdData.ReDraw = False    
	    
    ggoSpread.Source = frm1.vspdData	    
        
    ggoSpread.CopyRow   
    
    With frm1			
   
		frm1.vspdData.ReDraw = True    
       
	    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow   
	    .vspdData.Focus
    	Call SetActiveCell(frm1.vspdData,C_GroupLevel,frm1.vspdData.ActiveRow,"M","X","X")
    End With
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
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)  

    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    With frm1	
		
		.vspdData.ReDraw = False
		.vspdData.focus

	    ggoSpread.Source = .vspdData
        ggoSpread.InsertRow , iIntReqRows

		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)

		.vspdData.ReDraw = True
     
    End With    

    Set gActiveElement = document.activeElement 

	If Err.number = 0 Then
		FncInserRow = True
	End IF

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
    
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
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
    
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
'******************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtItemGroup=" & Trim(.hItemGroup.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		
    Else   
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtItemGroup=" & Trim(.txtItemGroup.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	
    End If
  
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()				'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetQuerySpreadColor()
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
	frm1.hItemGroup.value = Trim(frm1.txtItemGroup.value)
	
    lgBlnFlgChgValue = False   
	
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call SetToolbar("11001111001111")

End Function



'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal, strDel
   	Dim iColSep
   	Dim TmpBufferVal, TmpBufferDel
   	Dim iTotalStrVal, iTotalStrDel
   	Dim iValCnt, iDelCnt
	Dim starDate
	Dim finaDate
	
    DbSave = False                                                          '⊙: Processing is NG
    
       Call LayerShowHide(1)
		
    On Error Resume Next
                                                       '☜: Protect system from crashing
	With frm1
		 .txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = Parent.gColSep
    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'☜: 신규 
				
				strVal = ""
				
				strVal = strVal & "C" & iColSep & lRow & iColSep					'☜: C=Create				
                
                .vspdData.Col = C_GroupLevel 	
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ItemGroup
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_CostElmtCd	
                strVal = strVal & Trim(.vspdData.Text) &  iColSep                         
                
                .vspdData.Col = C_ComCostElmtCd	
                strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep                               
                
                ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = StrVal
                iValCnt = iValCnt + 1                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
            
				strVal = ""

				strVal = strVal & "U" & iColSep						'☜: U=Update
				strVal = strVal &lRow & iColSep	

                .vspdData.Col = C_GroupLevel 	
                strVal = strVal & Trim(.vspdData.Text) & iColSep                
                .vspdData.Col = C_ItemGroup	              
                strVal = strVal & Trim(.vspdData.Text) & iColSep        
                
                .vspdData.Col = C_CostElmtCd
                strVal = strVal & Trim(.vspdData.Text) &  iColSep	
                .vspdData.Col = C_ComCostElmtCd
                strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep	
                
				ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = StrVal
                iValCnt = iValCnt + 1                                                                                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag												'☜: 삭제 
            
				strDel = ""

				strDel = strDel & "D" & iColSep
				strDel = strDel & lRow & iColSep	

                .vspdData.Col = C_GroupLevel 
                strDel = strDel & Trim(.vspdData.Text) & iColSep                
                .vspdData.Col = C_ItemGroup 
                strDel = strDel & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_CostElmtCd 
                strDel = strDel & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ComCostElmtCd 
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferDel(iDelCnt)
                
                TmpBufferDel(iDelCnt) = StrDel
                iDelCnt = iDelCnt + 1 
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
	
	iTotalStrVal = Join(TmpBufferVal, "")
	iTotalStrDel = Join(TmpBufferDel, "")
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True																	'⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	Call MainQuery()

End Function


Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
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

'check item_group	
	If Trim(frm1.txtItemGroup.value) <> "" Then
		strWhere = " item_group_cd  = " & FilterVar(frm1.txtItemGroup.value, "''", "S") & " "		
		
		Call CommonQueryRs(" item_group_nm  ","	 ufn_c_get_itemgroup() ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtItemGroup.alt,"X")
			frm1.txtItemGroup.focus 
			frm1.txtItemGroupNM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtItemGroupNM.value = strDataNm(0)
	End If

End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원부자재그룹별원가요소등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>품목그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroup" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup 'CON',frm1.txtItemGroup.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%>> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
