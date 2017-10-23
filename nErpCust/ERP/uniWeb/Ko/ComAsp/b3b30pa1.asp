<%@ LANGUAGE="VBSCRIPT" %>
<!--
=========================================================================================================
'*  1. Module Name          : Production																*
'*  2. Function Name        : Popup Characteristic														*
'*  3. Program ID           : b3b30pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Characteristic Popup												        *
'*  7. Modified date(First) : 2003/02/03					 											*
'*  8. Modified date(Last)  : 																            *
'*  9. Modifier (First)     : Lee Woo Guen   															*
'* 10. Modifier (Last)      :       																    *
'* 11. Comment              :																			*
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../inc/incSvrCcm.inc" -->
<!-- #Include file="../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">  <!-- '��: �ش� ��ġ�� ���� �޶���, ��� ��� -->

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "b3b30pb1.asp"							 '��: �����Ͻ� ���� ASP�� 

Dim C_CharCd
Dim C_CharNm
Dim C_CharValueDigit

<!-- #Include file="../inc/lgVariables.inc" -->

Dim strReturn                                              '--- Return Parameter Group 

Dim lgCurDate
Dim lgIntConFlg
Dim IsOpenPop

Dim gblnWinEvent                                             '~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
                                                             'PopUp Window�� ��������� ���θ� ��Ÿ���� variable %>
Dim arrReturn
Dim arrParent
Dim arrParam					
Dim arrField
Dim PlantCd

Dim PopupParent
				
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_CharCd			= 1
	C_CharNm			= 2
	C_CharValueDigit	= 3
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										'��: Initializes Group View Size
	lgStrPrevKeyIndex = ""										'initializes Previous Key
	lgIntConFlg = 0
	lgIntFlgMode = PopupParent.OPMD_CMODE
	
    lgSortKey = 1                                       '��: initializes sort direction
	gblnWinEvent = False
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "PA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
	
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	txtCharCd.value = arrParam(0)	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim i

	Call InitSpreadPosVariables()

    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

    vspdData.ReDraw = False
	    
'    vspdData.OperationMode = 3

    vspdData.MaxCols = C_CharValueDigit + 1
    vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit C_CharCD,			"����׸�",		20
	ggoSpread.SSSetEdit C_CharNM,			"����׸��",	30
	ggoSpread.SSSetEdit C_CharValueDIGIT,	"��簪�ڸ���",	14,1

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)

	ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
	
	vspdData.ReDraw = True
	
	Call SetSpreadLock()
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
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
			C_CharCd			= iCurColumnPos(1)
			C_CharNm			= iCurColumnPos(2)
			C_CharValueDigit	= iCurColumnPos(3)
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
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' Ÿ��Ʋ cell�� dblclick�߰ų�....
	   Exit Function
	End If

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

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 and vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then
					Exit Sub
				End If
			End If
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub

'========================================================================================================
'	Name : OKClick()
'	Desc : handle ok icon click event
'========================================================================================================
Function OKClick()
	Dim i, iCurColumnPos
	
	If vspdData.MaxRows > 0 Then
		
		Redim arrReturn(UBound(arrField))

        ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
			
		For i = 0 To UBound(arrField)
			If arrField(i) <> "" Then
				vspddata.Col = iCurColumnPos(CInt(arrField(i)))
				arrReturn(i) = vspdData.Text
			End If
		Next
	
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
				
End Function

'========================================================================================================
'	Name : CancelClick()
'	Desc : handle  Cancel click event
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : MousePointer()
'	Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
	Call LoadInfTB19029																'��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	
	Call InitVariables

	Call InitComboBox()

	Call SetDefaultVal()
	
	Call InitSpreadSheet()
	
	If DbQuery = False Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'	Name : FncQuery()
'	Desc : 
'========================================================================================================
Function FncQuery()
	FncQuery = False
	Call InitVariables()
		
	If txtCharCd.value = "" And txtCharNm.value  <> "" Then
		lgIntConFlg = 1							'Code�� ��ȸ�ϴ� �� �̸����� ��ȸ�ϴ� �� ���� 
	End If
	
	vspdData.MaxRows = 0						'Grid �ʱ�ȭ 
	
	If DbQuery = False Then
		Exit Function
	End If
	
	FncQuery = True
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

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    'Err.Clear         
                                                          '��: Protect system from crashing
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
   
    DbQuery = False                                                         '��: Processing is NG
	
	Call LayerShowHide(1)												<%'��: �۾������� ǥ�� %>	
    
    
    Dim strVal
    
 	strVal = BIZ_PGM_ID & "?txtMode="		& PopupParent.UID_M0001												'��: 
	
	strVal = strVal & "&txtCharCd="			& Trim(txtCharCd.value)					'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtCharNm="			& txtCharNm.value
	
	strVal = strVal & "&lgIntConFlg="		& lgIntConFlg
	strVal = strVal & "&txtMaxRows="        & vspdData.MaxRows
	strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex   

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
	End If
    lgIntFlgMode = PopupParent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
End Function

</SCRIPT>
<!-- #Include file="../inc/Uni2kCMCom.inc" -->
</HEAD>
<%
'########################################################################################################
'#						6. Tag ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>
				<TR>
					<TD CLASS=TD5 NOWRAP>����׸�</TD>
					<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharCd" SIZE=25 MAXLENGTH=18 tag="11XXXU" ALT="����׸�"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP></TD>	
					<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtCharNm" SIZE=40 MAXLENGTH=40 tag="11" ALT="����׸��"></TD>
				</TR>		
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/b3b30pa1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK = "FncQuery()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>

</BODY>
</HTML>
