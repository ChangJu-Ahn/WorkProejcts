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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>계획차수</TITLE>
<!--
'########################################################################################################
'#						1. 선 언 부																		#
'########################################################################################################
-->
<!--
'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncServer.asp" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBS">
	Option Explicit					<% '☜: indicates that All variables must be declared in advance %>

<%
'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
%>
<%
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
%>

<%
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(GetSvrDate, gServerDateFormat, gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, gDateFormat)
%>

	Const BIZ_PGM_QRY_ID = "s2111pb3.asp"			'☆: 비지니스 로직 ASP명 

	Const C_PlanSeq = 1								'☆: Spread Sheet 의 Columns 인덱스 

	Const C_SHEETMAXROWS = 30

<%
'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
%>
	Dim arrReturn								'--- Return Parameter Group
	Dim lgIntGrpCount							'☜: Group View Size를 조사할 변수 

	Dim lgStrPrevKey
	Dim gblnWinEvent							'~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
												'	PopUp Window가 사용중인지 여부를 나타내는 variable

	Const lsPLANNUM  = "PLANNUM"				'계획차수 

<%
'============================================  1.2.3 Global Variable값 정의  ============================
'========================================================================================================
%>
<% '----------------  공통 Global 변수값 정의  ------------------------------------------------------- %>

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++ %>

<%
'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
%>
<% 
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
%>
<%
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
%>
	Function InitVariables()
		lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
		lgStrPrevKey = ""										<%'initializes Previous Key%>
		
		<% '------ Coding part ------ %>
		gblnWinEvent = False
		ReDim arrReturn(0)
		Self.Returnvalue = arrReturn
	End Function
	
<%
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
%>
<%
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
%>
	Sub SetDefaultVal()

		Dim arrTemp
		
		arrTemp = Split(window.dialogArguments, gColSep)

		txtPlanSeq.value = arrTemp(0)

		txtConSalesOrg.value = arrTemp(1)
		txtConSpYear.value = arrTemp(2)
		txtConPlanTypeCd.value = arrTemp(3)
		txtConDealTypeCd.value = arrTemp(4)
		txtConCurr.value = arrTemp(5)
		txtSelectChr.value = arrTemp(6)

	End Sub
<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
<% '== 조회,출력 == %>
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

		vspdData.MaxCols = C_PlanSeq
		vspdData.MaxRows = 0

		vspdData.OperationMode = 3

		vspdData.ReDraw = False

		ggoSpread.SpreadInit
		ggoSpread.SSSetEdit	C_PlanSeq, "계획차수", 18,2

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
'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	Function OKClick()
	
		If vspdData.ActiveRow > 0 Then	
			Redim arrReturn(vspdData.MaxCols - 1)
		
			vspdData.Row = vspdData.ActiveRow
			vspdData.Col = C_PlanSeq
			arrReturn(0) = vspdData.Text

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
'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
%>
<%
'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
%>
<%
'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
'========================================================================================================
%>
<%
'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
%>
	Sub Form_Load()
		Call LoadInfTB19029
		Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
		Call InitSpreadSheet()
		Call SetDefaultVal
		Call InitVariables
		Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")
		Call FncQuery()
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
'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
%>

<%
'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
%>
<%
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)
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
<%
'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
%>
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If
    
		If vspdData.MaxRows < NewTop + C_SHEETMAXROWS And lgStrPrevKey <> "" Then
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If	
					
				
				If DBQuery = False Then
					
					Exit Sub
				End If
		End if    


	End Sub

<%
'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
%>
<%
'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
%>

<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>

	Function FncQuery() 
	    
	    FncQuery = False                                                        <%'⊙: Processing is NG%>
	    
	    Err.Clear                                                               <%'☜: Protect system from crashing%>

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '⊙: Clear Contents  Field %>
		Call InitVariables													<% '⊙: Initializes local global variables %>

	<%  '-----------------------
	    'Query function call area
	    '----------------------- %>
	    Call DbQuery																<%'☜: Query db data%>

	    FncQuery = True																<%'⊙: Processing is OK%>
	        
	End Function


<%
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
%>
	Function DbQuery()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		Dim strVal
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & lsPLANNUM					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtConSalesOrg=" & Trim(txtConSalesOrg.value)	<%'☜: 조회 조건 데이타 %>
		strVal = strVal & "&txtConSpYear=" & Trim(txtConSpYear.value)
		strVal = strVal & "&txtConPlanTypeCd=" & Trim(txtConPlanTypeCd.value)
		strVal = strVal & "&txtConDealTypeCd=" & Trim(txtConDealTypeCd.value)
		strVal = strVal & "&txtConCurr=" & Trim(txtConCurr.value)
		strVal = strVal & "&txtSelectChr=" & Trim(txtSelectChr.value)

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbQuery = True														<%'⊙: Processing is NG%>
	End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
	<TABLE CELLSPACING=0 CLASS="basicTB">
<!--		<TR>
			<TD HEIGHT=40>
				<FIELDSET>
					<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
						<TR>
							<TD CLASS="TD5" NOWRAP>계획차수</TD>
							<TD CLASS="TD6"><INPUT NAME="txtPlanSeq" ALT="계획차수" TYPE="Text" MAXLENGTH=3 SiZE=10 tag="11"></TD>
						</TR>
					</TABLE>
				</FIELDSET>
			</TD>
		</TR>
-->		<TR>
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE WIDTH="100%" HEIGHT="100%">
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s2111pa3_I433441817_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=30>
				<TABLE CLASS="basicTB" CELLSPACING=0>
					<TR>
<!--						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()"  
						     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"></IMG></TD>
-->						<TD WIDTH=30% ALIGN=RIGHT>
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

	<INPUT TYPE=HIDDEN NAME="txtPlanSeq" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtConSalesOrg" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtConSpYear" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtConPlanTypeCd" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtConDealTypeCd" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtConCurr" tag="14">
	<INPUT TYPE=HIDDEN NAME="txtSelectChr" tag="14">

	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>