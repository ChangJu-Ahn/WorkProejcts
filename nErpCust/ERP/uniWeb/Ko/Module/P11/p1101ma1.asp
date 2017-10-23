
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1101ma1.asp
'*  4. Program Name         : 기간 생성 
'*  5. Program Desc         :
'*  6. Component List		:
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/18
'*  9. Modifier (First)     : Mr  Kim Gyoung-Don
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!--
========================================================================================================
=                          1.1.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--==========================================  1.1.2 공통 Include   ======================================
==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim BaseDate
Dim strYear
Dim strMonth
DIm strDay
Dim lgMaxYear
DIm lgMinYear

'========================================================================================================
'=                       1.2.1 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

BaseDate = "<%=GetSvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

lgMaxYear = StrYear + 20
lgMinYear = StrYear - 10

Const BIZ_PGM_BATCH_ID = "p1101mb2.asp"												
Const BIZ_PGM_LOOKUP_ID = "p1101mb4.asp"											

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               
    lgBlnFlgChgValue = False                                                
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														

End Sub

'=============================== 2.1.2 LoadInfTB19029() =================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	
	Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")
	Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
	
	frm1.txtDayCnt.text = ""
	frm1.txtYear.text = StrYear
	
End Sub

Sub InitComboBox()
	Call SetCombo(frm1.cboPeriodType, "01", "월")								'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
    Call SetCombo(frm1.cboPeriodType, "02", "순")
    Call SetCombo(frm1.cboPeriodType, "03", "주")
    Call SetCombo(frm1.cboPeriodType, "04", "일")
    
    Call SetCombo(frm1.cboWeekDay, "2", "월")								'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
    Call SetCombo(frm1.cboWeekDay, "3", "화")
    Call SetCombo(frm1.cboWeekDay, "4", "수")
    Call SetCombo(frm1.cboWeekDay, "5", "목")
    Call SetCombo(frm1.cboWeekDay, "6", "금")
    Call SetCombo(frm1.cboWeekDay, "7", "토")
    Call SetCombo(frm1.cboWeekDay, "1", "일")
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------

Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "칼렌다 타입 팝업"			' 팝업 명칭 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtClnrType.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "칼렌다 타입"					' TextBox 명칭 
	
    arrField(0) = "CAL_TYPE"						' Field명(0)
    arrField(1) = "CAL_TYPE_NM"						' Field명(1)
    
    arrHeader(0) = "칼렌다 타입"				' Header명(0)
    arrHeader(1) = "칼렌다 타입명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1) 
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'#########################################################################################################
'******************************************  3.1 Window 처리  ********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	
	Call AppendNumberPlace("6","4","0")
	Call AppendNumberRange("6",lgMinYear,lgMaxYear)
	Call AppendNumberPlace("7","3","0")
	Call AppendNumberRange("7","1","365")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
	frm1.txtClnrType.focus 
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
'   Event Name : cboPeriodType_OnChange()
'   Event Desc :
'==========================================================================================
Sub cboPeriodType_onChange()
	If frm1.cboPeriodType.value = "03" Then
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"N")
		frm1.cboWeekDay.value = "1"
		
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")		
		frm1.txtDayCnt.text = ""

	ElseIf frm1.cboPeriodType.value = "04" Then
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"N")		
		frm1.txtDayCnt.text = "1"
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
		frm1.cboWeekDay.value = ""	
	Else
		Call ggoOper.SetReqAttr(frm1.cboWeekDay,"Q")
		frm1.cboWeekDay.value = ""	
		Call ggoOper.SetReqAttr(frm1.txtDayCnt,"Q")		
		frm1.txtDayCnt.text = "" 		
	End If	
End Sub

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################

'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	Dim IntRetCD
		
	lgBlnFlgChgValue = False 
	
	If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
	 
	If Not chkField(Document, "2") Then
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) < CInt(lgMinYear) Then
		Call DisplayMsgBox("970023","X","년도",lgMinYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If CInt(frm1.txtYear.Value) > CInt(lgMaxYear) Then
		Call DisplayMsgBox("972004","X","년도",lgMaxYear)
		frm1.txtYear.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.cboPeriodType.value = "04" Then
		If frm1.txtDayCnt.Value < 1 Then
			Call DisplayMsgBox("970023","X","기간내일수","1")
			frm1.txtDayCnt.focus 
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
		
	Call LookUpLotPeriod	
	
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
    Call parent.FncExport(Parent.C_SINGLE)												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : 
'========================================================================================

Function FncCancel() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : LookUpLotPeriod
' Function Desc : 기간 생성 버튼을 누르면 생성된 기간이 있는 지 조회한다.
'========================================================================================

Function LookUpLotPeriod()

	Dim strVal
	
    LayerShowHide(1)
		
    With frm1
		.txtUpdtUserId.value = Parent.gUsrID
		
		strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
		strVal = strVal & "&txtYear=" & Trim(.txtYear.text)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtUpdtUserId=" & Trim(.txtUpdtUserId.value)
	
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

End Function

Function LotPerdLookUpOk()
	Dim rtnVal
	
	rtnVal = DisplayMsgBox("188100",Parent.VB_YES_NO,"X","X")
	
	If rtnVal = vbYes Then
		Call DbExecute
	Else
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	End If

End Function
'========================================================================================
' Function Name : DbExecute
' Function Desc : 실행 후 정상적으로 수행되었을 경우에 
'========================================================================================
Function DbExecute()
    
    With frm1
		.txtMode.value = Parent.UID_M0002											'☜: 저장 상태 
		'.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
		.txtInsrtUserId.value  = Parent.gUsrID
		.txtUpdtUserId.value = Parent.gUsrID
	End With
       
    Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										'☜: 비지니스 ASP 를 가동 
	
End Function

Function DbExecOk()
	Call DisplayMsgBox("183114","X","X","X")	
End Function


Function LotPerdNo()
	
		Call LayerShowHide(0)
		Call BtnDisabled(0)
	
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기간생성</font></td>
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
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>칼렌다 타입</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="22XXXU" ALT="칼렌다 타입"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 MAXLENGTH=30 tag="24" ALT="칼렌다 타입명"></TD>							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>년도</TD>
								<TD CLASS=TD6 NOWRAP>	
									<script language =javascript src='./js/p1101ma1_I169267725_txtYear.js'></script>								
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>기간범위</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPeriodType" ALT="기간범위" STYLE="Width: 98px;" tag="22"></SELECT></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>기간 시작요일</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboWeekDay" ALT="기간 시작요일" STYLE="Width: 98px;" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
    						<TR>
								<TD CLASS=TD5 NOWRAP>기간내일수</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p1101ma1_I603434699_txtDayCnt.js'></script>								
								</TD>
							</TR> 							
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" Flag=1 ONCLICK=FncSave>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
