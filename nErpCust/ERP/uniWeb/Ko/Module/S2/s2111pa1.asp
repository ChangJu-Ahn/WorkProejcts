<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : S/O Reference ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Cho Sung Hyun																*
'* 10. Modifier (Last)      : sonbumyeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>������</TITLE>
<!--
'########################################################################################################
'#						1. �� �� ��																		#
'########################################################################################################
-->
<!--
'********************************************  1.1 Inc ����  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncServer.asp" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<!--
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBS">
	Option Explicit					<% '��: indicates that All variables must be declared in advance %>

<%
'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
%>
<%
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
%>

<%
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(GetSvrDate, gServerDateFormat, gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, gDateFormat)
%>

<%
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
%>
	Dim arrReturn					<% '--- Return Parameter Group %>
	Dim cntAllCheck
<%
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
%>
<% '----------------  ���� Global ������ ����  ------------------------------------------------------- %>

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++ %>

<%
'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
%>
<% 
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
%>
<%
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
%>
	
<%
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
%>
<%
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
%>
	Sub SetDefaultVal()

		Dim strData, i

		ReDim arrReturn(0)
		Self.Returnvalue = arrReturn

		cntAllCheck = 1
		
'		top.document.title = Trim(window.dialogArguments) & "-" & window.document.title

		top.document.title = "����"

		For i = 1 To 12
			strData = strData & Chr(11)
			strData = strData & Chr(11) & i
			strData = strData & Chr(11) & Chr(12)
		Next	

		ggoSpread.Source = vspdData 
		ggoSpread.SSShowData strData

	End Sub
<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()

End Sub	
<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
%>
	Sub InitSpreadSheet()
		ggoSpread.Source = vspdData

		vspdData.MaxCols = 2
		vspdData.MaxRows = 0

		vspdData.ReDraw = False

		ggoSpread.SpreadInit
		ggoSpread.SSSetCheck 1,"", 2,,,True
		ggoSpread.SSSetEdit	2, "��", 8

		ggoSpread.SpreadLockWithOddEvenRowColor()
	
		vspdData.ReDraw = True

	End Sub

<%
'==========================================  2.2.4 SetSpreadLock()  =====================================
'=	Name : SetSpreadLock()																				=
'=	Description : This method set color and protect in spread sheet celles								=
'========================================================================================================
%>
<%
'==========================================  2.2.6 InitComboBox()  ======================================
'=	Name : InitComboBox()																				=
'=	Description : Combo Display																			=
'========================================================================================================
%>
<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	Function OKClick()
	
		Dim gRow, gVar

		Redim arrReturn(11)

		For gRow = 1 To 12
			vspdData.Row = gRow
			vspdData.Col = 1
			If vspdData.Text = "1" Then
				vspdData.Col = 2
				gVar = gVar + 1
				arrReturn(gVar) = vspdData.Text & gRowSep
			End If
		Next

		arrReturn(0) = gVar & gRowSep
		Self.Returnvalue = arrReturn

		Self.Close()
	End Function	
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
	Function CancelClick()
		Self.Close()
	End Function
<% 
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
%>
<%
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
%>
<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>
<%
'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
%>
	Sub Form_Load()
		Call LoadInfTB19029
		Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
		Call InitSpreadSheet()
		Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")
		Call SetDefaultVal
	End Sub
<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	End Sub
<%
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
%>

<%
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
%>
<%
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)
		
		If Row <> 0 Or Col <> 1 Then Exit Function

		Dim gRow

		Select Case cntAllCheck Mod 2
		Case 0
			For gRow = 1 To 12
				vspdData.Row = gRow
				vspdData.Col = 1
				vspdData.Text = "0"
			Next
		Case 1
			For gRow = 1 To 12
				vspdData.Row = gRow
				vspdData.Col = 1
				vspdData.Text = "1"
			Next
		End Select

		cntAllCheck = cntAllCheck + 1

	End Function


<%
'====================================  3.3.2 vspdData_ButtonClicked()  ==================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
%>
	Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

		With vspdData 
	
			ggoSpread.Source = vspdData
   
			If Row > 0 And Col = 1 Then
			    If ButtonDown = 0 Then Exit Sub
			    .Row = Row
				If .Text = "" Then .Text = "0"
			End If
    
		End With

	End Sub


<%
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
'=	Event Name : vspdData_LeaveCell																		=
'=	Event Desc :																						=
'========================================================================================================
%>
<%
'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
%>
<%
'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
%>
<%
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
%>

<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>
<%
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
%>
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
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE WIDTH="100%" HEIGHT="100%">
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s2111pa1_I247520154_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=30>
				<TABLE CLASS="basicTB" CELLSPACING=0>
					<TR>
						<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()"
						     onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" 
						     onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>

	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>