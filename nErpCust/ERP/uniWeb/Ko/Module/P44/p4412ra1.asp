<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: 																			*
'*  3. Program ID			: p4412ra1                           										*
'*  4. Program Name			: Reference Popup GI for Order List											*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2000/04/17																*
'*  8. Modified date(Last)	: 2002/12/12																*
'*  9. Modifier (First)     : Kim, Gyoung-Don															*
'* 10. Modifier (Last)		: Ryu Sung Won																*	
'* 11. Comment 		:	
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin																				*
'********************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->

<!--'********************************************  1.1 Inc ����  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 ���� Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

Const BIZ_PGM_ID = "p4412rb1.asp"

Dim C_CompntCd
Dim C_CompntNm
Dim C_CompntSpec
Dim C_IssueDt
Dim C_IssueQty
Dim C_Unit
Dim C_WcCd
Dim C_LotNo
Dim C_LotSubNo
Dim C_SlipNo
Dim C_MoveType
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgPlantCD
Dim lgProdOrderNo
		
'*********************************************  1.3 �� �� �� ��  ****************************************
'*	����: Constant�� �ݵ�� �빮�� ǥ��.																*
'********************************************************************************************************

Dim arrParent
Dim arrParam					
		
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCD = arrParent(1)
lgProdOrderNo = arrParent(2)

top.document.title = PopupParent.gActivePRAspName
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
	C_CompntCd	= 1
	C_CompntNm	= 2
	C_CompntSpec= 3
	C_IssueDt	= 4
	C_IssueQty	= 5
	C_Unit		= 6
	C_WcCd		= 7
	C_LotNo		= 8
	C_LotSubNo	= 9
	C_SlipNo	= 10
	C_MoveType	= 11
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount	= 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey	= ""                           'initializes Previous Key		
	lgStrPrevKey1	= ""
	lgStrPrevKey2	= ""
	lgIntFlgMode	= PopupParent.OPMD_CMODE
	Self.Returnvalue = Array("")
End Function
	
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
	    
    vspdData.ReDraw = False
	        
    vspdData.MaxCols = C_MoveType + 1
    vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit		C_CompntCd, "��ǰ��", 18
	ggoSpread.SSSetEdit		C_CompntNm, "��ǰ���", 25
	ggoSpread.SSSetEdit		C_CompntSpec,"�԰�", 25
	ggoSpread.SSSetDate		C_IssueDt,	"�����", 11, 2, PopupParent.gDateFormat
	ggoSpread.SSSetFloat	C_IssueQty, "������",15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_Unit,		"����", 8
	ggoSpread.SSSetEdit		C_WcCd,		"�۾���", 10
	ggoSpread.SSSetEdit		C_LotNo,	"Lot ��ȣ", 12
	ggoSpread.SSSetEdit		C_LotSubNo, "����", 8
	ggoSpread.SSSetEdit		C_SlipNo,	"����ȣ", 18
	ggoSpread.SSSetEdit		C_MoveType, "�����", 10

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	
	ggoSpread.SSSetSplit2(2)							'frozen ����߰� 
		
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
			C_CompntCd	= iCurColumnPos(1)
			C_CompntNm	= iCurColumnPos(2)
			C_CompntSpec= iCurColumnPos(3)
			C_IssueDt	= iCurColumnPos(4)
			C_IssueQty	= iCurColumnPos(5)
			C_Unit		= iCurColumnPos(6)
			C_WcCd		= iCurColumnPos(7)
			C_LotNo		= iCurColumnPos(8)
			C_LotSubNo	= iCurColumnPos(9)
			C_SlipNo	= iCurColumnPos(10)
			C_MoveType	= iCurColumnPos(11)
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

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"
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

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=27 Then
 		Call CancelClick()		
	End If
End Sub	


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
	Call InitVariables
	Call ggoOper.LockField(Document, "N")                       '��: Lock  Suitable  Field
		
	Call SetDefaultVal()
	Call InitSpreadSheet()
	
	If DbQuery = False Then	
			Exit Sub
	End If
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
     On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
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
    Err.Clear								'��: Protect system from crashing
	    
    DbQuery = False							'��: Processing is NG
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal =  BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&lgPlantCD=" & lgPlantCD							'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&lgProdOrderNo=" & lgProdOrderNo					'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
    strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2	    
	    
    Call RunMyBizASP(MyBizASP, strVal)				<%'��: �����Ͻ� ASP �� ���� %>

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
	
    lgIntFlgMode = PopupParent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
						<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrdNo" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="����������ȣ"></TD>
						<TD CLASS=TD5 NOWRAP>ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/p4412ra1_vspdData_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
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
