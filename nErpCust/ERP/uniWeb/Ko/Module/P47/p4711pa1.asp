
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Operation Popup															*
'*  3. Program ID			: p4711pa1.asp																*
'*  4. Program Name			: �̷¹�ȣ Popup															*
'*  5. Program Desc			: �̷¹�ȣ Popup															*
'*  7. Modified date(First)	: 2001/12/14																*
'*  8. Modified date(Last)	: 2002/12/11																*
'*  9. Modifier (First)     : Park, Bum-Soo																*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 				:
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin																			*
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

Option Explicit
	
'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
Const BIZ_PGM_QRY_ID = "p4711pb1.asp"					<% '��: �����Ͻ� ���� ASP�� %>
		
Dim C_BatchRunNo
Dim C_ExecStartDt
Dim C_ProdtOrderNoFrom
Dim C_ProdtOrderNoTo
Dim C_ItemCdFrom
Dim C_ItemCdTo
Dim C_WcCdFrom
Dim C_WcCdTo
Dim C_ShiftCdFrom
Dim C_ShiftCdTo
Dim C_ReportDtFrom
Dim C_ReportDtTo
Dim C_Status
Dim C_Success_Cnt
Dim C_Error_Cnt
Dim C_InsrtUserId
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim arrReturn
Dim IsOpenPop
Dim lgNextNo
Dim lgPrevNo
Dim lgPlantCD
Dim ArrParent

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

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
	C_BatchRunNo		= 1
	C_ExecStartDt		= 2
	C_ProdtOrderNoFrom	= 3
	C_ProdtOrderNoTo	= 4
	C_ItemCdFrom		= 5
	C_ItemCdTo			= 6
	C_WcCdFrom			= 7
	C_WcCdTo			= 8
	C_ShiftCdFrom		= 9
	C_ShiftCdTo			= 10
	C_ReportDtFrom		= 11
	C_ReportDtTo		= 12
	C_Status			= 13
	C_Success_Cnt		= 14
	C_Error_Cnt			= 15
	C_InsrtUserId		= 16
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey = ""                           'initializes Previous Key		
	Self.Returnvalue = Array("")
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub
	
Function InitSetting()
	txtPlantCd.Value = ArrParent(1)
	txtPlantNm.Value = ArrParent(2)
	txtBatchRunNo.Value = ArrParent(3)
End Function
	
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
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_InsrtUserId + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit		C_BatchRunNo,			"�̷¹�ȣ", 18
	ggoSpread.SSSetDate 	C_ExecStartDt,			"������", 12, 2, gDateFormat
	ggoSpread.SSSetEdit		C_ProdtOrderNoFrom,		"���ۿ�����ȣ", 18
	ggoSpread.SSSetEdit		C_ProdtOrderNoTo,		"���������ȣ", 18
	ggoSpread.SSSetEdit		C_ItemCdFrom,			"����ǰ��", 18
	ggoSpread.SSSetEdit		C_ItemCdTo,				"����ǰ��", 18
	ggoSpread.SSSetEdit		C_WcCdFrom,				"�����۾���", 10
	ggoSpread.SSSetEdit		C_WcCdTo,				"�����۾���", 10
	ggoSpread.SSSetEdit		C_ShiftCdFrom,			"���� Shift", 10
	ggoSpread.SSSetEdit		C_ShiftCdTo,			"���� Shift", 10
	ggoSpread.SSSetDate 	C_ReportDtFrom,			"���۽�����", 12, 2, gDateFormat
	ggoSpread.SSSetDate 	C_ReportDtTo,			"���������", 12, 2, gDateFormat
	ggoSpread.SSSetEdit		C_Success_Cnt,			"���ȵȽ�����", 10
	ggoSpread.SSSetEdit		C_Error_Cnt,			"������", 10
	ggoSpread.SSSetEdit		C_InsrtUserId,			"������ID", 13
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_Status,C_Status, True)
    
    ggoSpread.SSSetSplit2(2)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub
	

'================================== 2.2.4 SetSpreadLock() ==================================================
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
			C_BatchRunNo		= iCurColumnPos(1)
			C_ExecStartDt		= iCurColumnPos(2)
			C_ProdtOrderNoFrom	= iCurColumnPos(3)
			C_ProdtOrderNoTo	= iCurColumnPos(4)
			C_ItemCdFrom		= iCurColumnPos(5)
			C_ItemCdTo			= iCurColumnPos(6)
			C_WcCdFrom			= iCurColumnPos(7)
			C_WcCdTo			= iCurColumnPos(8)
			C_ShiftCdFrom		= iCurColumnPos(9)
			C_ShiftCdTo			= iCurColumnPos(10)
			C_ReportDtFrom		= iCurColumnPos(11)
			C_ReportDtTo		= iCurColumnPos(12)
			C_Status			= iCurColumnPos(13)
			C_Success_Cnt		= iCurColumnPos(14)
			C_Error_Cnt			= iCurColumnPos(15)
			C_InsrtUserId		= iCurColumnPos(16)
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
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
		Redim arrReturn(3)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_BatchRunNo
			arrReturn(0) = vspdData.Text
			vspdData.Col = C_Status
			arrReturn(1) = vspdData.Text
			vspdData.Col = C_Success_Cnt
			arrReturn(2) = vspdData.Text
			vspdData.Col = C_Error_Cnt
			arrReturn(3) = vspdData.Text
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
	Call InitSetting()
	Call FncQuery()
	vspdData.Row = 1
	vspdData.Col = 1
	Call SetFocusToDocument("M")
	vspdData.Focus
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
	
    Err.Clear                                                               <%'��: Protect system from crashing%>
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then									'��: This function check indispensable field
	   Exit Function
	End If
	    
    DbQuery = False                                                         <%'��: Processing is NG%>
	    
    vspdData.MaxRows = 0
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001				<%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
    strVal = strVal & "&txtPlantCd=" & txtPlantCd.Value						<%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtBatchRunNo=" & txtBatchRunNo.Value
	If rdoDeleteFlg1.checked = True Then
		strVal = strVal & "&txtrdoflag=" & "C"
	Else
		strVal = strVal & "&txtrdoflag=" & "R"
	End If

    Call RunMyBizASP(MyBizASP, strVal)					<%'��: �����Ͻ� ASP �� ���� %>
		
    DbQuery = True                                                          			<%'��: Processing is NG%>

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = PopupParent.OPMD_UMODE													'��: Indicates that current mode is Update mode    
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TD CLASS=TD5 NOWRAP>����</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14XXXU" ALT="����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�̷¹�ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchRunNo" SIZE=18 MAXLENGTH=18 tag="11XXXU"  ALT="�̷¹�ȣ"></TD>
						<TD CLASS=TD5 NOWRAP>��ұ���</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDeleteFlg" ID="rdoDeleteFlg1" CLASS="RADIO" tag="11" Value="Y"><LABEL FOR="rdoDeleteFlg1">��</LABEL>
						     				 <INPUT TYPE="RADIO" NAME="rdoDeleteFlg" ID="rdoDeleteFlg2" CLASS="RADIO" tag="11" Value="N" CHECKED><LABEL FOR="rdoDeleteFlg2">�ƴϿ�</LABEL></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4711pa1_vspdData_vspdData.js'></script>
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
