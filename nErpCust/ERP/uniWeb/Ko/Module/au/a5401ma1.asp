
<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a5401ma1
'*  4. Program Name         : 미결관리기준등록 
'*  5. Program Desc         : 미결관리기준등록 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/11/7
'*  8. Modified date(Last)  : 2002/11/7
'*  9. Modifier (First)     : Jung Sung Ki
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'*
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'********************************************  1.1 Inc 선언   ********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--

'============================================  1.1.1 Style Sheet  =======================================
'======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->
<!--
'============================================  1.1.2 공통 Include  ======================================
'======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance 


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ====================================
'==========================================================================================================

Const BIZ_PGM_ID = "a5401mb1.asp"											 '☆: 비지니스 로직 ASP명 

'============================================  1.2.2 Global 변수 선언  ===================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2. Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'☆: 사용자 변수 초기화 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
	if Trim(frm1.txtAcctBaseNo.value)="" then
		frm1.txtAcctBaseNo.value = frm1.hAcctBaseNo.value
	end if

	frm1.txtAcctBaseNo.focus
	Set gActiveElement = document.activeElement
  
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'--------------------------------------------------------------------------------------------------------- 

Sub InitComboBox_One()
	Dim IntRetCD1
	Dim IntValMM
	Dim IntValDD
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F5004", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCardMM ,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F5005", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCardDD ,lgF0  ,lgF1  ,Chr(11))

    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub





Function OpenAcctBaseNo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "미결관리팝업"					' 팝업 명칭 
	arrParam(1) = "A_OPEN_ACCT_BASE"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "미결관리코드"
	
    arrField(0) = "ACCT_BASE_NO"							' Field명(0)
    arrField(1) = "ACCT_BASE_NM"						' Field명(1)
    
    arrHeader(0) = "미결관리"					' Header명(0)
    arrHeader(1) = "미결관리명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	frm1.txtAcctBaseNo.focus
	    Exit Function
	Else
		frm1.txtAcctBaseNo.focus
		frm1.txtAcctBaseNo.value = arrRet(0)
		frm1.txtAcctBaseNm.value = arrRet(1)
	End If	

End Function



'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'Sub Combo_Change(Index As Integer)
'	lgBlnFlgChgValue = True
'End Sub


'###########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################


'==========================================================================================================
Sub Form_Load()

    Call InitVariables																'⊙: Initializes local global variables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("1100100000001111")
    Call InitComboBox_One

	frm1.txtAcctBaseNo.focus 
	frm1.txtAcctBaseNo.value="1"

    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed

	FncQuery 

	Set gActiveElement = document.activeElement


End Sub


'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

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


'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing
  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables															'⊙: Initializes local global variables
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'☜: Query db data
    FncQuery = True																'⊙: Processing is OK
        
End Function



'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    Call SetToolbar("1100100000001111")

	frm1.txtAcctBaseNo.focus

    FncNew = True																'⊙: Processing is OK
    Set gActiveElement = document.activeElement
    
End Function


'========================================================================================

Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														'⊙: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003",Parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    

	FncSave = False                                                         '⊙: Processing is NG

	Err.Clear                                                               '☜: Protect system from crashing
	    
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If

   

	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then 		                                     '☜: Save db data 
		Exit Function
	End If    
	    
	FncSave = True                                                          '⊙: Processing is OK
    
End Function



'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
   
End Function



'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function



'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function



'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function



'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    parent.FncPrint()
End Function


'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing

End Function


'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing

End Function


'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'☜: 화면 유형 
End Function


'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================

Function DbDelete() 
    On Error Resume Next                                                    '☜: Protect system from crashing

End Function


'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function




'========================================================================================
' Function Name : cboXCH_RATE_FG_OnChange
' Function Desc : 
'========================================================================================

Sub txtCashAmt_Change() 
	lgBlnFlgChgValue = True
End Sub

Sub cboCardMM_OnChange() 
	lgBlnFlgChgValue = True
End Sub


Sub cboCardDD_OnChange()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================

Function DbQuery() 
    
    Err.Clear                                                               '☜: Protect system from crashing
    DbQuery = False                                                         '⊙: Processing is NG
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing
    

    Dim strVal
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtAcctBaseNo=" & Trim(frm1.txtAcctBaseNo.value)				'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG

End Function


'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("1100100000011111")
   
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
End Function



'========================================================================================

Function DbSave() 

    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

    Dim strVal
    
    Call LayerShowHide(1)                                                   '☜: Protect system from crashing

	With frm1
	
		.txtMode.value = Parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value     = lgIntFlgMode
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With

    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================

Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
    
    lgBlnFlgChgValue = False
    
    FncQuery

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>미결관리등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>미결관리기준</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAcctBaseNo" MAXLENGTH="2" SIZE=10 ALT ="미결관리기준" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenAcctBaseNo(frm1.txtAcctBaseNo.value,0)">&nbsp;
													<INPUT NAME="txtAcctBaseNm" MAXLENGTH="30" SIZE=30 ALT ="미결관리기준명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>

						<TABLE  <%=LR_SPACE_TYPE_60%>>
							<TR >
								<TD CLASS="TD5" NOWRAP HEIGHT=30>현금지급조건</TD>
							    <TD CLASS="TD6" NOWRAP COLSPAN="3"><script language =javascript src='./js/a5401ma1_I492656892_txtCashAmt.js'></script>&nbsp; 원 이하 </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP HEIGHT=30>신용카드지급조건</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCardMM" ALT="신용카드지급조건" STYLE="WIDTH: 100px" tag="22"></SELECT> 달전<SELECT NAME="cboCardDD" ALT="신용카드지급조건" STYLE="WIDTH: 100px" tag="22"></SELECT> 일까지</TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
								<TR HEIGHT="*">
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6 COLSPAN="3">&nbsp;</TD>
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>	
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hAcctBaseNo" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

