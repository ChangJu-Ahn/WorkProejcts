<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name		: Production
'*  2. Function Name	: 
'*  3. Program ID		: b1b01ma2.asp
'*  4. Program Name		: ǰ�������ȸ 
'*  5. Program Desc		:
'*  6. Component List	: 
'*  7. Modified date(First)	: 2000/12/18
'*  8. Modified date(Last)	: 2002/11/14
'*  9. Modifier (First)		: Jung Yu Kyung
'* 10. Modifier (Last)		: Hong Chang Ho
'* 11. Comment		:
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../Inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID			= "b1b02mb3.asp"		
Const BIZ_PGM_JUMPITEM_ID		= "b1b01ma1"
Const BIZ_PGM_JUMPITEMIMAGE_ID	= "b1b02ma1"

Const DIR_INIT_FILE = "../../../CShared/image/unierp20logo.gif"

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim C_Item
Dim C_ItmNm
Dim C_ItmFormalNm
Dim C_ItmAcc
Dim C_Unit
Dim C_ItmGroupCd
Dim C_ItmGroupNm
Dim C_Phantom
Dim C_BlanketPur
Dim C_BaseItm
Dim C_BaseItmNm
Dim C_SumItmClass
Dim C_DefaultFlg
Dim C_PicFlg
Dim C_ItmSpec
Dim C_UnitWeight
Dim C_UnitOfWeight
Dim C_GrossWeight
Dim C_UnitOfGrossWeight
Dim C_CBM
Dim C_CBMDesc
Dim C_DrawNo
Dim C_HsCd
Dim C_HsUnit
Dim C_StartDt
Dim C_EndDt

Dim lgOldRow
Dim IsOpenPop
Dim lgStrPrevKey1
Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Item				= 1
	C_ItmNm				= 2
	C_ItmFormalNm		= 3
	C_ItmAcc			= 4
	C_Unit				= 5
	C_ItmGroupCd		= 6
	C_ItmGroupNm		= 7
	C_Phantom			= 8
	C_BlanketPur		= 9
	C_BaseItm			= 10
	C_BaseItmNm			= 11
	C_SumItmClass		= 12
	C_DefaultFlg		= 13
	C_PicFlg			= 14
	C_ItmSpec			= 15
	C_UnitWeight		= 16
	C_UnitOfWeight		= 17
	C_GrossWeight		= 18
	C_UnitOfGrossWeight = 19
	C_CBM				= 20
	C_CBMDesc			= 21
	C_DrawNo			= 22
	C_HsCd				= 23
	C_HsUnit			= 24
	C_StartDt			= 25
	C_EndDt				= 26
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			
    lgBlnFlgChgValue = False			
    lgIntGrpCount = 0					
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
	lgStrPrevKey1 = ""
    lgSortKey = 1                                       '��: initializes sort direction
	lgOldRow = 0
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "MA")%>
End Sub

'=============================== 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_EndDt + 1										'��: �ִ� Columns
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_Item, "ǰ��",	15
		ggoSpread.SSSetEdit 	C_ItmNm, "ǰ���", 30
		ggoSpread.SSSetEdit 	C_ItmFormalNm, "ǰ�����ĸ�", 30
		ggoSpread.SSSetEdit 	C_ItmAcc, "ǰ�����", 15
		ggoSpread.SSSetEdit 	C_Unit, "����",	10
		ggoSpread.SSSetEdit 	C_ItmGroupCd, "ǰ��׷�", 15
		ggoSpread.SSSetEdit 	C_ItmGroupNm, "ǰ��׷��", 15
		ggoSpread.SSSetEdit 	C_Phantom, "����", 15, 2
		ggoSpread.SSSetEdit 	C_BlanketPur, "���ձ���", 15, 2
		ggoSpread.SSSetEdit 	C_BaseItm, "����ǰ��", 15
		ggoSpread.SSSetEdit 	C_BaseItmNm, "����ǰ���", 30
		ggoSpread.SSSetEdit 	C_SumItmClass, "�����ǰ��Ŭ����", 15
		ggoSpread.SSSetEdit 	C_DefaultFlg, "��ȿ����", 15, 2
		ggoSpread.SSSetEdit 	C_PicFlg, "��������", 15, 2
		ggoSpread.SSSetEdit 	C_ItmSpec, "ǰ��԰�", 15
		ggoSpread.SSSetFloat	C_UnitWeight,"Net�߷�", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetEdit 	C_UnitOfWeight, "Net����", 10
		ggoSpread.SSSetFloat	C_GrossWeight,	 "Gross�߷�",15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_UnitOfGrossWeight, "Gross����",10
		ggoSpread.SSSetFloat	C_CBM, "CBM(����)",15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit 	C_CBMDesc, "CBM����", 20
		ggoSpread.SSSetEdit 	C_DrawNo, "�����ȣ", 15
		ggoSpread.SSSetEdit 	C_HsCd, "HS�ڵ�", 15
		ggoSpread.SSSetEdit 	C_HsUnit, "HS����", 10
		ggoSpread.SSSetDate		C_StartDt, "������", 12, 2, parent.gDateFormat
		ggoSpread.SSSetDate		C_EndDt, "������", 12, 2, parent.gDateFormat
	
		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 

		Call ggoSpread.SSSetColHidden(C_BaseItmNm, C_BaseItmNm, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'=========================== 2.2.4 SetSpreadLock() ======================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================= 2.2.5 SetSpreadColor() =======================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.SSSetRequired C_Item,	pvStartRow, pvEndRow

    .vspdData.ReDraw = True
    
    End With
End Sub

Sub InitComboBox()
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1002", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemClass, lgF0, lgF1, Chr(11))
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
			C_Item				= iCurColumnPos(1)
			C_ItmNm				= iCurColumnPos(2)
			C_ItmFormalNm		= iCurColumnPos(3)
			C_ItmAcc			= iCurColumnPos(4)
			C_Unit				= iCurColumnPos(5)
			C_ItmGroupCd		= iCurColumnPos(6)
			C_ItmGroupNm		= iCurColumnPos(7)
			C_Phantom			= iCurColumnPos(8)
			C_BlanketPur		= iCurColumnPos(9)
			C_BaseItm			= iCurColumnPos(10)
			C_BaseItmNm			= iCurColumnPos(11)
			C_SumItmClass		= iCurColumnPos(12)
			C_DefaultFlg		= iCurColumnPos(13)
			C_PicFlg			= iCurColumnPos(14)
			C_ItmSpec			= iCurColumnPos(15)
			C_UnitWeight		= iCurColumnPos(16)
			C_UnitOfWeight		= iCurColumnPos(17)
			C_GrossWeight		= iCurColumnPos(18) 
			C_UnitOfGrossWeight	= iCurColumnPos(19)
			C_CBM				= iCurColumnPos(20) 
			C_CBMDesc			= iCurColumnPos(21)
			C_DrawNo			= iCurColumnPos(22)
			C_HsCd				= iCurColumnPos(23)
			C_HsUnit			= iCurColumnPos(24)
			C_StartDt			= iCurColumnPos(25)
			C_EndDt				= iCurColumnPos(26)
    End Select    
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================  2.2.1 SetDefaultVal()  ==================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'===================================================================================================
Sub SetDefaultVal()
	frm1.txtFinishStartDt.Text	=  StartDate
	frm1.txtFinishEndDt.Text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
		
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtHighItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtHighItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "ǰ��׷�"			
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtHighItemGroupCd.focus
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	frm1.hItemCd.value = arrRet(2)
	
End Function

'------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtHighItemGroupCd.Value    = arrRet(0)		
	frm1.txtHighItemGroupNm.Value    = arrRet(1)		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function JumpItem()
	
	Dim IntRetCd, strVal, Row,Col
	
    Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_Item
	
	If Row <= 0 Then
		WriteCookie "txtItemCd", UCase(Trim(frm1.txtItemCd.value))
		WriteCookie "txtItemNm", frm1.txtItemNm.value 
	Else
		WriteCookie "txtItemCd", UCase(Trim(frm1.vspdData.Text))
		
		frm1.vspdData.Col = C_ItmNm
		WriteCookie "txtItemNm", UCase(Trim(frm1.vspdData.Text))
	End If
	
	PgmJump(BIZ_PGM_JUMPITEM_ID)
	
End Function

Function JumpItemImage()
	
	Dim IntRetCd, strVal, Row,Col
	
    Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_Item
	
	If Row <= 0 Then
		WriteCookie "txtItemCd", UCase(Trim(frm1.txtItemCd.value))
		WriteCookie "txtItemNm", frm1.txtItemNm.value 
	Else
		WriteCookie "txtItemCd", UCase(Trim(frm1.vspdData.Text))
		
	End If
	
	PgmJump(BIZ_PGM_JUMPITEMIMAGE_ID)
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	
    Call LoadInfTB19029	  															'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
   
	Call SetDefaultVal
   	Call InitComboBox
    Call InitVariables		
    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet	
	Call SetToolbar("11000000000011")
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement 

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
'	If Row <= 0 Or Col < 0 Then
'		Exit Sub
'	End If
	
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000111111")
	
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If

	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = C_Item
		frm1.vspdData.Row = Row
		lgOldRow = Row
		
		Call DbDtlQuery(frm1.vspdData.Text)
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 1.����ð�(runtime)�� �˾��޴��� ���ؼ� �������� �ٲ���.
'				 2.Mouse�� Ư��Cell�� ����("SPC")�ϰ� ������ ��ư("SPCR")�� ������ �˾��� ���δ�.
'				   �˾����� Ư�� �޴� item�� ����("SPCRP") ���� Į���� freeze�Ѵ�.
'=======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
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
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFinishStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFinishStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFinishStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFinishEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFinishEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFinishEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFinishStartDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFinishEndDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                       
    
    Err.Clear                                                              
    
	'-----------------------
    'Erase contents area
    '----------------------- 

	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtHighItemGroupCd.value = "" Then
		frm1.txtHighItemGroupNm.value = ""
	End If
	
	If ValidDateCheck(frm1.txtFinishStartDt, frm1.txtFinishEndDt) = False Then
		Exit Function
	End If

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables														
        
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
	If DbQuery = False Then   
		Exit Function           
    End If 
           
    FncQuery = True														
        
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
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
   Dim strAvailableItem
	
	Err.Clear															

	DbQuery = False														

	LayerShowHide(1)
		
	Dim strVal
	
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)		
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.hItemAcct.value)
		strVal = strVal & "&cboItemClass=" & Trim(frm1.hSumItemClass.value)
		strVal = strVal & "&txtHighItemGroupCd=" & Trim(frm1.hItemGroup.value)
		strVal = strVal & "&txtFinishStartDt=" & Trim(frm1.hStartDt.value)
		strVal = strVal & "&txtFinishEndDt=" & Trim(frm1.hEndDt.value)
		strVal = strVal & "&rdoDefaultFlg=" & frm1.hAvailableItem.value	
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.cboItemAcct.value)
		strVal = strVal & "&cboItemClass=" & Trim(frm1.cboItemClass.value)
		strVal = strVal & "&txtHighItemGroupCd=" & Trim(frm1.txtHighItemGroupCd.value)
		strVal = strVal & "&txtFinishStartDt=" & Trim(frm1.txtFinishStartDt.text)
		strVal = strVal & "&txtFinishEndDt=" & Trim(frm1.txtFinishEndDt.text)
		If frm1.rdoDefaultFlg1.checked = True then
			strAvailableItem = "A"
		ElseIf frm1.rdoDefaultFlg2.checked = True then
			strAvailableItem = "Y"
		Else
			strAvailableItem = "N"
		End IF
		strVal = strVal & "&rdoDefaultFlg=" & strAvailableItem
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If	
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True																				
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()													
	Call ggoOper.LockField(Document, "Q")								
	Call SetToolbar("11000000000111")
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call vspdData_Click(1,1)
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
	lgIntFlgMode = parent.OPMD_UMODE										
End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryNotOk()													
	document.all.ImgItemImage.src= DIR_INIT_FILE
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(strItemCd) 
	
	Dim strVal
	
	Err.Clear															

	DbDtlQuery = False													
	
	If CommonQueryRs(" ITEM_CD "," b_item_image "," ITEM_CD = " & FilterVar(strItemCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then   
		Call DisplayMsgBox("122900","X","X","X")
		document.all.ImgItemImage.src= DIR_INIT_FILE
		Exit Function
	End If
		
	strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & strItemCd		  '��: query key
	strVal = strVal     & "&txtDKeyValue=" & "default"                            '��: default value
	strVal = strVal     & "&txtTable="     & "b_item_image"                       '��: Table Name
	strVal = strVal     & "&txtField="     & "item_image"	                      '��: Field
	strVal = strVal     & "&txtKey="       & "item_cd"	                          '��: Key
	
	document.all.ImgItemImage.src = ValueEscape(strVal)
		
	DbDtlQuery = True																					
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�������ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����� ǰ��Ŭ����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="����� ǰ��Ŭ����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btHighItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/b1b02ma2_txtFinishStartDt_txtFinishStartDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/b1b02ma2_txtFinishEndDt_txtFinishEndDt.js'></script>
	
									</TD>
									<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg1" CLASS="RADIO" tag="1X" Value="ALL" CHECKED><LABEL FOR="rdoDefaultFlg1">��ü</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg2" CLASS="RADIO" tag="1X" Value="Y"><LABEL FOR="rdoDefaultFlg2">��</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoDefaultFlg" ID="rdoDefaultFlg3" CLASS="RADIO" tag="1X" Value="N"><LABEL FOR="rdoDefaultFlg3">�ƴϿ�</LABEL>
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
								<TD HEIGHT=* WIDTH=50%>
									<script language =javascript src='./js/b1b02ma2_I757896786_vspdData.js'></script>
								</TD>								
								<TD HEIGHT=* WIDTH=50%>
									<IFRAME NAME="ImgItemImage" SRC="../../../CShared/image/unierp20logo.gif" marginwidth=0 marginheight=0 WIDTH=100% HEIGHT=100% FRAMEBORDER=1 FRAMESPACING=0></IFRAME>
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
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItem">ǰ����</A>&nbsp;|&nbsp;<A href="vbscript:JumpItemImage">ǰ��������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hSumItemClass" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroup" tag="24"><INPUT TYPE=HIDDEN NAME="hAvailableItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
