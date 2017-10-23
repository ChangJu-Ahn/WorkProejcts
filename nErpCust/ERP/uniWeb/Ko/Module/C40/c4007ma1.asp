<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name			: ���������� 
'*  2. Function Name		: 
'*  3. Program ID			: c4007ma1.asp
'*  4. Program Name			: ��������׷캰������ҵ�� 
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

Option Explicit                                                             '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_ID = "c4007mb1.asp"			'��: Head Query �����Ͻ� ���� ASP�� 

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

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop						' Popup
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

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
	lgSortKey = 1
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
       
		ggoSpread.SSSetEdit		C_GroupLevel,	"�׷�Level", 12,,,5,2
		ggoSpread.SSSetButton 	C_GroupLevelPopup 
		ggoSpread.SSSetEdit		C_ItemGroup,	"ǰ��׷�", 10,,,10  
		ggoSpread.SSSetButton		C_ItemGroupPopup 
		ggoSpread.SSSetEdit		C_ItemGroupNM,		"ǰ��׷��", 25
		ggoSpread.SSSetEdit 		C_CostElmtCd,	"�������1",12,,,10  
		ggoSpread.SSSetButton 	C_CostElmtPopup
		ggoSpread.SSSetEdit		C_CostElmtNM,			"������Ҹ�1", 20
		ggoSpread.SSSetEdit 		C_ComCostElmtCd,	"�������2",12,,,10  
		ggoSpread.SSSetButton 	C_ComCostElmtPopup
		ggoSpread.SSSetEdit		C_ComCostElmtNM,			"������Ҹ�2", 20
		
		Call ggoSpread.MakePairsColumn(C_GroupLevel, C_GroupLevelPopup)
		Call ggoSpread.MakePairsColumn(C_ItemGroup, C_ItemGroupPopup )
		Call ggoSpread.MakePairsColumn(C_CostElmtCd, C_CostElmtPopup )
		Call ggoSpread.MakePairsColumn(C_ComCostElmtCd, C_ComCostElmtPopup )
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		ggoSpread.SSSetSplit2(4)										'frozen ����߰� 
				
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
' Function Desc : �׸��� ��� Ŭ���� ���� 
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
' Function Desc : �׸��� �����¸� �����Ѵ�.
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

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
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
			arrParam(0) = "�׷�Level"						' �˾� ��Ī 
			arrParam(1) = " ( SELECT DISTINCT GROUP_LEVEL FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE ��Ī 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�׷�Level"							' TextBox ��Ī	
	
			arrField(0) = "ED12" & Parent.gColSep & "group_level"					' Field��(1)
			     
			arrHeader(0) = "�׷�Level"						' Header��(0)

		
		Case C_CostElmtPopup 
			arrParam(0) = "�������1�˾�"						' �˾� ��Ī 
			arrParam(1) = " C_COST_ELMT_S  "			' TABLE ��Ī 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = " cost_elmt_type='M' "			' Where Condition
			arrParam(5) = "�������"							' TextBox ��Ī 
	
			arrField(0) = "ED10" & Parent.gColSep & "cost_elmt_cd"					' Field��(1)
			arrField(1) = "ED25" & Parent.gColSep & "cost_elmt_nm"					' Field��(0)
			     
			arrHeader(0) = "�������"						' Header��(0)
			arrHeader(1) = "������Ҹ�"						' Header��(0)   			

		Case C_ComCostElmtPopup 
			arrParam(0) = "�������2�˾�"						' �˾� ��Ī 
			arrParam(1) = " C_COST_ELMT_S  "			' TABLE ��Ī 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = " cost_elmt_type='M' "			' Where Condition
			arrParam(5) = "�������"							' TextBox ��Ī 
	
			arrField(0) = "ED10" & Parent.gColSep & "cost_elmt_cd"					' Field��(1)
			arrField(1) = "ED25" & Parent.gColSep & "cost_elmt_nm"					' Field��(0)
			     
			arrHeader(0) = "�������"						' Header��(0)
			arrHeader(1) = "������Ҹ�"						' Header��(0)   		
		Case C_ItemGroupPopup
			frm1.vspdData.Col = C_GroupLevel  : 			 frm1.vspdData.Row = frm1.vspdData.ActiveRow
			

			arrParam(0) = "ǰ��׷��˾�"						' �˾� ��Ī 
			arrParam(1) = " UFN_C_GET_ITEMGROUP() "			' TABLE ��Ī 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			If Trim(frm1.vspdData.Text) <>"" Then
			arrParam(4) = " group_level = "			& filterVar(	Trim(frm1.vspdData.Text),"''","S") 	& " "		' Where Condition
			Else
			arrParam(4) = ""							' Where Condition
			End If
			arrParam(5) = "ǰ��׷�"							' TextBox ��Ī 
	
	
			arrField(0) = "HH" & Parent.gColSep & "item_group_cd"					' Field��(0)
			arrField(1) = "ED12" & Parent.gColSep & "group_level"					' Field��(1)
			arrField(2) = "ED12" & Parent.gColSep & "item_group_cd"					' Field��(0)
			arrField(3) = "ED20" & Parent.gColSep & "item_group_nm"					' Field��(1)

			     
			arrHeader(0) = "ǰ��׷�"						' Header��(0)
			arrHeader(1) = "�׷�Level"						' Header��(1)
			arrHeader(2) = "ǰ��׷�"						' Header��(0)
			arrHeader(3) = "ǰ��׷��"						' Header��(1)
			
		Case Else
			arrParam(0) = "ǰ��׷��˾�"						' �˾� ��Ī 
			arrParam(1) = " UFN_C_GET_ITEMGROUP() "			' TABLE ��Ī 
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "ǰ��׷�"							' TextBox ��Ī 
	
	
			arrField(0) = "HH" & Parent.gColSep & "item_group_cd"					' Field��(0)
			arrField(1) = "ED12" & Parent.gColSep & "group_level"					' Field��(1)
			arrField(2) = "ED12" & Parent.gColSep & "item_group_cd"					' Field��(0)
			arrField(3) = "ED20" & Parent.gColSep & "item_group_nm"					' Field��(1)

			     
			arrHeader(0) = "ǰ��׷�"						' Header��(0)
			arrHeader(1) = "�׷�Level"						' Header��(1)
			arrHeader(2) = "ǰ��׷�"						' Header��(0)
			arrHeader(3) = "ǰ��׷��"						' Header��(1)
					
		
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
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetItemPopup()  --------------------------------------------------
'	Name : SetItemPopup()
'	Description : OpenItemPopup Popup���� Return�Ǵ� �� setting
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
			iStrFromList=" ( SELECT DISTINCT GROUP_LEVEL FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE ��Ī 
			iStrWhereList = " group_level =" &  filtervar(pvStrData, "''","S")
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","�׷�Level","X")
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
			iStrFromList=" ( SELECT DISTINCT GROUP_LEVEL, ITEM_GROUP_CD, ITEM_GROUP_NM,  UPPER_ITEM_GROUP_CD FROM UFN_C_GET_ITEMGROUP() ) AA"			' TABLE ��Ī 
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
			iStrFromList=" C_COST_ELMT_S "			' TABLE ��Ī 
			iStrWhereList = " cost_elmt_cd =" &  filtervar(pvStrData, "''","S")
			iStrWhereList = iStrWhereList & " and cost_elmt_type='M'"
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","�������1","X")
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
			iStrFromList=" C_COST_ELMT_S "			' TABLE ��Ī 
			iStrWhereList = " cost_elmt_cd =" &  filtervar(pvStrData, "''","S")
			iStrWhereList = iStrWhereList & " and cost_elmt_type='M'"
			Call CommonQueryRs(iStrSelectList,iStrFromList , iStrWhereList,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			IF Len(lgF0) < 1 Then 
				Call DisplayMsgBox("970000","X","�������2","X")
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
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field

	Call InitSpreadSheet                                                    '��: Setup the Spread sheet
	Call InitVariables                                                      '��: Initializes local global variables

	
	'----------  Coding part  -------------------------------------------------------------	
	Call SetToolbar("11001111001111")										'��: ��ư ���� ����	
   
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop)	Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
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
	
    FncQuery = False															'��: Processing is NG

    Err.Clear																    '��: Protect system from crashing
	
	IF ChkKeyField()=False Then Exit Function 
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
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
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
    Call InitVariables
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     												'��: Query db data

    FncQuery = True																'��: Processing is OK

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
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    
    If Not ggoSpread.SSDefaultCheck Then              '��: Check required field(Multi area)
		Exit Function
    End If  
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then									'��: This function check indispensable field
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
    End If     																				'��: Save db data
    
    FncSave = True																'��: Processing is OK
           
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
	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)  

    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

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
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
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
    ' �����Ͱ� ���� ��� 
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
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
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
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()												'��: ���� ������ ���� ���� 
	
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
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
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtItemGroup=" & Trim(.hItemGroup.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		
    Else   
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtItemGroup=" & Trim(.txtItemGroup.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	
    End If
  
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()				'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetQuerySpreadColor()
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
	frm1.hItemGroup.value = Trim(frm1.txtItemGroup.value)
	
    lgBlnFlgChgValue = False   
	
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
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
	
    DbSave = False                                                          '��: Processing is NG
    
       Call LayerShowHide(1)
		
    On Error Resume Next
                                                       '��: Protect system from crashing
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

            Case ggoSpread.InsertFlag												'��: �ű� 
				
				strVal = ""
				
				strVal = strVal & "C" & iColSep & lRow & iColSep					'��: C=Create				
                
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

				strVal = strVal & "U" & iColSep						'��: U=Update
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
                
            Case ggoSpread.DeleteFlag												'��: ���� 
            
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
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True																	'��: Processing is NG

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 

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
'       					6. Tag�� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������׷캰������ҵ��</font></td>
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
									<TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroup" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup 'CON',frm1.txtItemGroup.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 tag="14"></TD>									
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
