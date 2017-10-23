<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211pa3.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Local L/C No POPUP ASP													*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2000/04/10																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'********************************************************************************************************
Response.Expires = -1													'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
%>
<HTML>
<HEAD>
<TITLE>EXPORT LOCAL L/C POPUP</TITLE>
<%
'########################################################################################################
'#						1. �� �� ��																		#
'########################################################################################################
%>
<%
'********************************************  1.1 Inc ����  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<%
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================
%>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/eventpopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/incImage.js"></SCRIPT>

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


	Const BIZ_PGM_QRY_ID = "m3211pb3.asp"			<% '��: �����Ͻ� ���� ASP�� %>

	Const C_LCNo = 1								<% '��: Spread Sheet �� Columns �ε��� %>
	Const C_LCDocNo = 2
	Const C_LCAmendSeq = 3
	Const C_OpenDt = 4
	Const C_ExpiryDt = 5
	Const C_AdvBank = 6
	Const C_LCType = 7

	Const C_SHEETMAXROWS = 30

<%
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
%>
	Dim strReturn					<% '--- Return Parameter Group %>
	Dim lgIntGrpCount				<% '��: Group View Size�� ������ ���� %>
	Dim arrReturn					<% '--- Return Parameter Group %>
	Dim lgStrPrevKey
	Dim gblnWinEvent				<% '~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
											   '	PopUp Window�� ��������� ���θ� ��Ÿ���� variable %>

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
	Function InitVariables()
		lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
		lgStrPrevKey = ""										<%'initializes Previous Key%>
		
		<% '------ Coding part ------ %>
		gblnWinEvent = False
		Self.Returnvalue = ""
	End Function
	
<%
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
%>

<%
'==========================================  2.2.2 LoadInfTB19029()  ====================================
'=	Name : LoadInfTB19029()																				=
'=	Description :  This method loads format inf															=
'========================================================================================================
%>
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/ComLoadInfTB19029.asp" -->		
End Sub
	
<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
%>
	Sub InitSpreadSheet()
		ggoSpread.Source = vspdData
		vspdData.OperationMode = 3

		vspdData.MaxCols = C_LCType
		vspdData.MaxRows = 0

		vspdData.ReDraw = False

		ggoSpread.SpreadInit

		ggoSpread.SSSetEdit		C_LCNo, "L/C������ȣ", 18, 0
		ggoSpread.SSSetEdit		C_LCDocNo, "L/C��ȣ", 20, 0
		ggoSpread.SSSetFloat	C_LCAmendSeq, "AMEND����", 15,	ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec,2,,"Z","0x","99x"
		ggoSpread.SSSetDate		C_OpenDt, "L/C������", 12, 2, gDateFormat
		ggoSpread.SSSetDate		C_ExpiryDt, "��ȿ��", 12, 2, gDateFormat
		ggoSpread.SSSetEdit		C_AdvBank, "�߽��Ƿ�����", 12, 0
		ggoSpread.SSSetEdit		C_LCType, "LOCAL L/C����", 12, 0

		SetSpreadLock "", 0, -1, ""

		vspdData.ReDraw = True
	End Sub
	
<%
'==========================================  2.2.4 SetSpreadLock()  =====================================
'=	Name : SetSpreadLock()																				=
'=	Description : This method set color and protect in spread sheet celles								=
'========================================================================================================
%>
	Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
		ggoSpread.Source = vspdData
			
		vspdData.ReDraw = False
			
		ggoSpread.SpreadLock C_LCNo, lRow, C_LCNo
		ggoSpread.SpreadLock C_LCDocNo, lRow, C_LCDocNo
		ggoSpread.SpreadLock C_LCAmendSeq, lRow, C_LCAmendSeq
		ggoSpread.SpreadLock C_OpenDt, lRow, C_OpenDt
		ggoSpread.SpreadLock C_ExpiryDt, lRow, C_ExpiryDt
		ggoSpread.SpreadLock C_AdvBank, lRow, C_AdvBank
		ggoSpread.SpreadLock C_LCType, lRow, C_LCType
			
		vspdData.ReDraw = True
	End Sub

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
		Dim intColCnt
		
		If vspdData.ActiveRow > 0 Then	
			Redim arrReturn(vspdData.MaxCols - 1)
		
			vspdData.Row = vspdData.ActiveRow
					
			For intColCnt = 0 To vspdData.MaxCols - 1
				vspdData.Col = intColCnt + 1
				arrReturn(intColCnt) = vspdData.Text
			Next
				
			Self.Returnvalue = arrReturn
		End If
		
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
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
%>
	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
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
'++++++++++++++++++++++++++++++++++++++++++++  OpenBizPartner()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenBizPartner()																				+
'+	Description : Business Partner PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenBizPartner()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function

		gblnWinEvent = True

		arrParam(0) = "������û��"							<%' �˾� ��Ī %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(txtApplicant.value)		<%' Code Condition%>
'		arrParam(3) = Trim(txtApplicantNm.value)		<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "������û��"								<%' TextBox ��Ī %>

		arrField(0) = "BP_CD"								<%' Field��(0)%>
		arrField(1) = "BP_NM"								<%' Field��(1)%>

		arrHeader(0) = "������û��"							<%' Header��(0)%>
		arrHeader(1) = "������û�θ�"						<%' Header��(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBizPartner(arrRet)
		End If
	End Function

<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>

<%
'+++++++++++++++++++++++++++++++++++++++++++  SetBizPartner()  ++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetBizPartner()																				+
'+	Description : Set Return array from Business Partner PopUp Window									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetBizPartner(arrRet)
		txtApplicant.Value = arrRet(0)
		txtApplicantNm.Value = arrRet(1)
	End Function
	
<%
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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

		<% ' �̹��� ȿ�� �ڹٽ�ũ��Ʈ �Լ� ȣ��  %>
		Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")

	
		Call LoadInfTB19029
		Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
		Call InitVariables
		Call InitSpreadSheet()
		'Call DbQuery()
		
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
'=====================================  3.2.2 btnApplicantOnClick()  ===================================
'========================================================================================================
%>
	Sub btnApplicantOnClick()
		Call OpenBizPartner()
	End Sub
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
		If Row = 0 Or vspdData.MaxRows = 0 Then 
          Exit Function
	    End If
		If vspdData.MaxRows > 0 Then
			If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End Function
	
<%
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
'=	Event Name : vspdData_LeaveCell																		=
'=	Event Desc :																						=
'========================================================================================================
%>

	Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
		With vspdData
			If Row >= NewRow Then
				Exit Sub
			End If

			If NewRow = .MaxRows Then
				If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
					DbQuery
				End If
			End If
		End With
	End Sub
	
<%
'========================================  3.3.3 vspdData_LeaveCell()  ==================================
'=	Event Name : vspdData_LeaveCell																		=
'=	Event Desc :																						=
'========================================================================================================
%>
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		If OldLeft <> NewLeft Then
			Exit Sub
		End If

		If NewTop > oldTop Then
			If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				DbQuery
			End If
		End If
	End Sub

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
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
%>
	Function DbQuery()
		Dim strVal
			    
		Err.Clear															<%'��: Protect system from crashing%>

		DbQuery = False														<%'��: Processing is NG%>

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '��: Clear Contents  Field %>
		Call InitVariables													<% '��: Initializes local global variables %>
		Call InitSpreadSheet()

		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then							<% '��: This function check indispensable field %>
			Exit Function
		End If

		if LayerShowHide(1) =false then
		    exit Function
		end if

	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & UID_M0001							<%'��: �����Ͻ� ó�� ASP�� ���� %>
	    strVal = strVal & "&txtApplicant=" & Trim(txtApplicant.value)		<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

		Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
		
	    DbQuery = True                   	
	End Function
	
Function DbQueryOk()
	lgIntFlgMode = OPMD_UMODE
	vspdData.Focus
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
			<FIELDSET >
				<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
					<TR>
						<TD CLASS=TD5>������û��</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="12XXXU" ALT="������û��"><IMG SRC="../../image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" onclick="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5></TD>
						<TD CLASS=TD6></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE WIDTH="100%" HEIGHT="100%" CLASS="TB3">
				<TR>
					<TD HEIGHT="100%"><script language =javascript src='./js/m3211pa3_vaSpread_vspdData.js'></script></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
</BODY>
</HTML>