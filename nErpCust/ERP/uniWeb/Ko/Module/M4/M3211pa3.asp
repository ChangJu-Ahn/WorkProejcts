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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
Response.Expires = -1													'☜ : ASP가 캐쉬되지 않도록 한다.
%>
<HTML>
<HEAD>
<TITLE>EXPORT LOCAL L/C POPUP</TITLE>
<%
'########################################################################################################
'#						1. 선 언 부																		#
'########################################################################################################
%>
<%
'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<%
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
%>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/eventpopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/incImage.js"></SCRIPT>

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


	Const BIZ_PGM_QRY_ID = "m3211pb3.asp"			<% '☆: 비지니스 로직 ASP명 %>

	Const C_LCNo = 1								<% '☆: Spread Sheet 의 Columns 인덱스 %>
	Const C_LCDocNo = 2
	Const C_LCAmendSeq = 3
	Const C_OpenDt = 4
	Const C_ExpiryDt = 5
	Const C_AdvBank = 6
	Const C_LCType = 7

	Const C_SHEETMAXROWS = 30

<%
'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
%>
	Dim strReturn					<% '--- Return Parameter Group %>
	Dim lgIntGrpCount				<% '☜: Group View Size를 조사할 변수 %>
	Dim arrReturn					<% '--- Return Parameter Group %>
	Dim lgStrPrevKey
	Dim gblnWinEvent				<% '~~~ ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
											   '	PopUp Window가 사용중인지 여부를 나타내는 variable %>

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
		Self.Returnvalue = ""
	End Function
	
<%
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
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

		ggoSpread.SSSetEdit		C_LCNo, "L/C관리번호", 18, 0
		ggoSpread.SSSetEdit		C_LCDocNo, "L/C번호", 20, 0
		ggoSpread.SSSetFloat	C_LCAmendSeq, "AMEND차수", 15,	ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec,2,,"Z","0x","99x"
		ggoSpread.SSSetDate		C_OpenDt, "L/C개설일", 12, 2, gDateFormat
		ggoSpread.SSSetDate		C_ExpiryDt, "유효일", 12, 2, gDateFormat
		ggoSpread.SSSetEdit		C_AdvBank, "추심의뢰은행", 12, 0
		ggoSpread.SSSetEdit		C_LCType, "LOCAL L/C유형", 12, 0

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
'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
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

		arrParam(0) = "개설신청인"							<%' 팝업 명칭 %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(txtApplicant.value)		<%' Code Condition%>
'		arrParam(3) = Trim(txtApplicantNm.value)		<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "개설신청인"								<%' TextBox 명칭 %>

		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>

		arrHeader(0) = "개설신청인"							<%' Header명(0)%>
		arrHeader(1) = "개설신청인명"						<%' Header명(1)%>

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
'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
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
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개별 프로그램마다 필요한 개발자 정의 Procedure(Sub, Function, Validation & Calulation 관련 함수)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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

		<% ' 이미지 효과 자바스크립트 함수 호출  %>
		Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")

	
		Call LoadInfTB19029
		Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
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
'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
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
				If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
			If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				DbQuery
			End If
		End If
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
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
%>
	Function DbQuery()
		Dim strVal
			    
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '⊙: Clear Contents  Field %>
		Call InitVariables													<% '⊙: Initializes local global variables %>
		Call InitSpreadSheet()

		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then							<% '⊙: This function check indispensable field %>
			Exit Function
		End If

		if LayerShowHide(1) =false then
		    exit Function
		end if

	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
	    strVal = strVal & "&txtApplicant=" & Trim(txtApplicant.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

		Call RunMyBizASP(MyBizASP, strVal)										<%'☜: 비지니스 ASP 를 가동 %>
		
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
'#						6. TAG 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET >
				<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
					<TR>
						<TD CLASS=TD5>개설신청인</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="12XXXU" ALT="개설신청인"><IMG SRC="../../image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" onclick="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
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