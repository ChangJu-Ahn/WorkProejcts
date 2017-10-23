<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC101PA1
'*  4. Program Name         : Delivery Item Popup Item by Plant 
'*  5. Program Desc         : Delivery Item Popup Item by Plant 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/22
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : 2003/02/22
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

Const BIZ_PGM_ID = "mc101pb1.asp"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_ItemCode    
Dim C_ItemName
Dim C_Spec
Dim C_Unit
Dim C_ItemAcctDesc
Dim C_ItemGroupCd
Dim C_ProcurTypeDesc
Dim C_LotFlg
Dim C_MajorSlCd
Dim C_IssuedSlCd
Dim C_ValidFlg
Dim C_RecvInspecFlg    

Dim IsOpenPop          
Dim arrReturn
Dim arrParent
Dim arrParam1
Dim arrParam2
Dim PlantCd

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)

arrParam1		= arrParent(1)
arrParam2		= arrParent(2)

top.document.title = PopupParent.gActivePRAspName

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgStrPrevKeyIndex	= ""
	lgLngCurRows		= 0
	lgSortKey			= 1
	Redim arrReturn(0)
    Self.Returnvalue	= arrReturn  
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","PA") %>
End Sub

'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
	On Error Resume Next
    Err.Clear
    
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemAccount,lgF0  ,lgF1  ,Chr(11))
End Sub
	
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	PlantCd					= arrParam1(0)
	frm1.txtItemCd.value	= arrParam1(1)
End Sub 

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030108", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 5
	    .MaxCols = C_RecvInspecFlg+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit  C_ItemCode,		"ǰ��",			20, 0, -1, 18
	    ggoSpread.SSSetEdit  C_ItemName,		"ǰ���",		25, 0, -1, 40
	    ggoSpread.SSSetEdit  C_Spec,			"�԰�",			25, 0, -1, 50
	    ggoSpread.SSSetEdit	 C_Unit,			"���ش���",		10, 0, -1, 3
	    ggoSpread.SSSetEdit  C_ItemAcctDesc,	"ǰ�����",		18, 0, -1, 50
	    ggoSpread.SSSetEdit  C_ItemGroupCd,		"ǰ��׷�",		18, 0, -1, 10
	    ggoSpread.SSSetEdit  C_ProcurTypeDesc,	"���ޱ���",		18, 0, -1, 50
	    ggoSpread.SSSetEdit  C_LotFlg,			"LOT����",		12, 0, -1, 1
	    ggoSpread.SSSetEdit  C_MajorSlCd,		"�԰�â��",		12, 0, -1, 7
	    ggoSpread.SSSetEdit  C_IssuedSlCd,		"���â��",		12, 0, -1, 7
	    ggoSpread.SSSetEdit  C_ValidFlg,		"��ȿ����",		12, 0, -1, 1
	    ggoSpread.SSSetEdit	 C_RecvInspecFlg,	"���԰˻籸��",	14, 0, -1, 1
	    
	    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    		
		ggoSpread.SSSetSplit(2)
		
		Call SetSpreadLock() 
		.ReDraw = true
    End With
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCode				= 1											'��: Spread Sheet�� Column�� ��� 
	C_ItemName				= 2
	C_Spec					= 3
	C_Unit					= 4
	C_ItemAcctDesc			= 5
	C_ItemGroupCd			= 6
	C_ProcurTypeDesc		= 7
	C_LotFlg				= 8
	C_MajorSlCd				= 9
	C_IssuedSlCd			= 10
	C_ValidFlg				= 11
	C_RecvInspecFlg			= 12
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
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
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_ItemCode			= iCurColumnPos(1)
    	C_ItemName			= iCurColumnPos(2)
	    C_Spec				= iCurColumnPos(3)
	    C_Unit				= iCurColumnPos(4)
	    C_ItemAcctDesc		= iCurColumnPos(5)
	    C_ItemGroupCd		= iCurColumnPos(6)
	    C_ProcurTypeDesc	= iCurColumnPos(7)
	    C_LotFlg			= iCurColumnPos(8)
	    C_MajorSlCd			= iCurColumnPos(9)
	    C_IssuedSlCd		= iCurColumnPos(10)
	    C_ValidFlg			= iCurColumnPos(11)
	    C_RecvInspecFlg		= iCurColumnPos(12)
	End Select
End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal lRow)
    frm1.vspdData.ReDraw = False
		
		ggoSpread.SSSetProtected	C_ItemCode,			lRow, lRow
		ggoSpread.SSSetProtected	C_ItemName,			lRow, lRow
		ggoSpread.SSSetProtected	C_Spec,				lRow, lRow
		ggoSpread.SSSetProtected	C_Unit,				lRow, lRow
		ggoSpread.SSSetProtected	C_ItemAcctDesc,		lRow, lRow
		ggoSpread.SSSetProtected	C_ItemGroupCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_ProcurTypeDesc,	lRow, lRow
		ggoSpread.SSSetProtected	C_LotFlg,			lRow, lRow
		ggoSpread.SSSetProtected	C_MajorSlCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_IssuedSlCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_ValidFlg,			lRow, lRow
		ggoSpread.SSSetProtected	C_RecvInspecFlg,	lRow, lRow
	
    frm1.vspdData.ReDraw = True
End Sub

'===========================================  2.3.1 ()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
  Dim intColCnt
  
  If frm1.vspdData.ActiveRow > 0 Then 
	Redim arrReturn(frm1.vspdData.MaxCols - 1)
  
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
	For intColCnt = 0 To frm1.vspdData.MaxCols - 1
		frm1.vspdData.Col			= intColCnt + 1
		arrReturn(intColCnt)		= frm1.vspdData.Text
	Next
    
	Self.Returnvalue = arrReturn
  End If
  
  Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()		
	Self.Close()
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    
'    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
    
    
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    
	'----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet
    
    Call InitVariables                                                      '��: Initializes local global variables
    Call InitComboBox()
    Call SetDefaultVal()
    Call InitSpreadSheet()
    
    If DbQuery = False Then
		Exit Sub
	End if
End Sub

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	End if
	
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing   

    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
	frm1.vspdData.Maxrows = 0
    Call InitVariables() 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then									'��: This function check indispensable field
'       Exit Function
'    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    																     '��: Query db data
    If DbQuery() = False Then
		Exit Function
	End if
       
    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Call LayerShowHide(1)
    
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
    Dim strVal
    
    strVal = BIZ_PGM_ID	& "?PlantCd="			& Trim(PlantCd)
    strVal = strVal		& "&txtItemCd="			& Trim(frm1.txtItemCd.value)
    strVal = strVal		& "&txtItemNm="			& Trim(frm1.txtItemNm.value)
    strVal = strVal		& "&cboItemAccount="	& Trim(frm1.cboItemAccount.value)
    strVal = strVal		& "&txtSpec="			& Trim(frm1.txtSpec.value)
    strVal = strVal     & "&txtMaxRows="		& frm1.vspdData.MaxRows
	strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	frm1.vspdData.Focus
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### 
-->

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="GET">


<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP>ǰ��</TD>
				<TD CLASS="TD6" COLSPAN = 3 NOWRAP><INPUT TYPE="Text" Name="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="11XXXU" ALT="ǰ���" TABINDEX="-1"></TD>
				
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>ǰ�����</TD>
				<TD CLASS="TD6" NOWRAP><SELECT NAME="cboItemAccount" ALT="ǰ�����" STYLE="Width: 160px;" tag="11" TABINDEX="-1"><OPTION VALUE=""></SELECT></TD>
				<TD CLASS="TD5" NOWRAP>�԰�</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtSpec" SIZE=50 MAXLENGTH=50 tag="11XXXU" ALT="�԰�" TABINDEX="-1"></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/mc101pa1_OBJECT1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
	
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
	
				<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="okclick()"    ></IMG>
						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


