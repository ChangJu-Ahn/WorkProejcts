
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Operation Popup															*
'*  3. Program ID			: p4112pa1.asp																*
'*  4. Program Name			: ���� Popup																*
'*  5. Program Desc			: ���� Popup																*
'*  7. Modified date(First)	: 2000/04/24																*
'*  8. Modified date(Last)	: 2002/12/10																*
'*  9. Modifier (First)     : Kim, Gyoung-Don															*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin				:																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--#####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->
<!--********************************************  1.1 Inc ����  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 ���� Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit					'��: indicates that All variables must be declared in advance

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
Const BIZ_PGM_QRY_ID = "p4112pb1_KO441.asp"					<% '��: �����Ͻ� ���� ASP�� %>
		
Dim C_OprNo
Dim C_JobCD
Dim C_JobDesc
Dim C_PlanStartDt
Dim C_PlanComptDt
Dim C_OrderStatus
Dim C_OrderStatusDesc
Dim C_InsideFlgHard
		
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim arrReturn				' Return Parameter Group
Dim IsOpenPop				' Popup

Dim lgPlantCD				' Plant Code
Dim lgInsideFlag
Dim ArrParent

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
lgPlantCD = ArrParent(1)
'txtProdOrderNo.Value = ArrParent(2)	'From_Load ���� ó�� 
lgInsideFlag = ArrParent(3)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_OprNo				= 1
	C_JobCD				= 2
	C_JobDesc			= 3
	C_PlanStartDt		= 4
	C_PlanComptDt		= 5
	C_OrderStatus		= 6
	C_OrderStatusDesc	= 7
	C_InsideFlgHard		= 8
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgStrPrevKey = ""                           'initializes Previous Key
	Self.Returnvalue = Array("")
End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox�� Value�� Setting�Ѵ�.														=
'========================================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    'Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'ggoSpread.Source = vspdData
    'ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
    'ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobDesc

    'Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'ggoSpread.Source = vspdData
    'ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    'ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex

	With vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_OrderStatus
			intIndex = .value
			.Col = C_OrderStatusDesc
			.value = intindex
			
		'	.Row = intRow
		'	.col = C_JobCd
		'	intIndex = .value
		'	.Col = C_JobDesc
		'	.value = intindex
		Next	
	End With
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20081125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_InsideFlgHard + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit 	C_OprNo,		"����", 10								' Operation No.
	ggoSpread.SSSetEdit 	C_JobCd,		"�۾���", 10								' Job Code
	ggoSpread.SSSetEdit 	C_JobDesc,		"�۾����", 15							' Job Description		
	ggoSpread.SSSetDate 	C_PlanStartDt,	"����������", 12, 2, parent.gDateFormat	' Planned Start Date
	ggoSpread.SSSetDate 	C_PlanComptDt,	"�ϷΌ����", 12, 2, parent.gDateFormat	' Planned Completion Date
	ggoSpread.SSSetCombo 	C_OrderStatus,	"���û���", 10					' Order Status
	ggoSpread.SSSetCombo 	C_OrderStatusDesc,	"���û���", 10				' Order Status
	ggoSpread.SSSetEdit 	C_InsideFlgHard,	"�系�ܱ���", 10					' Order Status

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_OrderStatus,C_OrderStatus, True)
	
	ggoSpread.SSSetSplit2(1)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_OprNo				= iCurColumnPos(1)
			C_JobCD				= iCurColumnPos(2)
			C_JobDesc			= iCurColumnPos(3)
			C_PlanStartDt		= iCurColumnPos(4)
			C_PlanComptDt		= iCurColumnPos(5)
			C_OrderStatus		= iCurColumnPos(6)
			C_OrderStatusDesc	= iCurColumnPos(7)
			C_InsideFlgHard		= iCurColumnPos(8)
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
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1,1)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then Exit Sub
		End If
    End if
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	If vspdData.MaxRows > 0 Then
		Dim intRowCnt
		Dim intColCnt
		Dim intSelCnt

		intSelCnt = 0
		Redim arrReturn(0)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_OprNo
			arrReturn(0) = vspdData.Text
		End If

		Self.Returnvalue = arrReturn
		Self.Close()
	End If			
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.activeRow > 0 Then
 		Call OkClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	

'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = lgPlantCD
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = Trim(txtProdOrderNo.value) 

	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
	
	iCalledAspName = AskPRAspName("p4111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
		
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtProdOrderNo.Focus
	
End Function

'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    txtProdOrderNo.Value    = arrRet(0)		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call InitVariables
	Call InitSpreadSheet()
	Call InitComboBox()
	txtProdOrderNo.Value = ArrParent(2)
	Call FncQuery()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************

'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
   FncQuery = False
	Call DbQuery()
   Fncquery = False
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################

'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then
	   Exit Function
	End If
	    
    DbQuery = False
	    
    vspdData.MaxRows = 0
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal & "&lgPlantCd=" & lgPlantCD
    strVal = strVal & "&txtProdOrderNo=" & txtProdOrderNo.Value	    
    strVal = strVal & "&lgInsideFlag=" & lgInsideFlag

    Call RunMyBizASP(MyBizASP, strVal)
		
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(lngStartRow)												'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    lgIntFlgMode = PopupParent.OPMD_UMODE													'��: Indicates that current mode is Update mode    
    Call InitData(lngStartRow,1)
    vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5>����������ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4112pa1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK = "FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
